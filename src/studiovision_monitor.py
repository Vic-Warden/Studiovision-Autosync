"""
studiovision_monitor.py
Image Router — latest monitor build targeting Windows workstations.

This module bridges an ophthalmic imaging device and the Studiovision patient
management software (Microsoft Access).  When a new medical image is written to
the watched source folder, the router:

  1. Waits until the file is fully written (lock-polling).
  2. Polls the currently active Access form to identify the patient on screen.
  3. Resolves that patient's dedicated photo folder via a query on PUBLIC.MDB.
  4. Moves the image into the patient folder.
  5. Inserts a matching record in the Documents table of PUBLIC.MDB.
  6. Triggers an immediate Requery/Refresh of the Access form so the clinician
     can see the newly linked document without any manual action.

Files whose patient cannot be resolved within PATIENT_WAIT_TIMEOUT seconds are
quarantined in ORPHAN_DIR for later review instead of being silently lost.

Architecture:
  Producer (main thread)  — Watchdog observer converts filesystem events into
                             Path objects and places them on a thread-safe queue.
  Consumer (worker thread) — Drains the queue sequentially, performing all
                              I/O-bound and COM-bound operations on a single
                              dedicated thread to avoid COM apartment conflicts.

Dependencies:
  watchdog  — cross-platform filesystem event monitoring
  pyodbc    — ODBC bridge to Microsoft Access (.mdb) databases
  pywin32   — COM automation client to interact with a running Access instance
  pythoncom — COM initialisation per-thread (required by pywin32)
"""

import pythoncom
import queue
import shutil
import sys
import threading
import time
import logging
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False

# Configuration 
SOURCE_DIR  = Path(r"??")
ORPHAN_DIR  = Path(r"??")
DEST_PHOTOS = Path(r"??")
PUBLIC_MDB  = Path(r"??")
DOCUM_MDB   = Path(r"??")

# Supported image extensions
WATCHED_EXTENSIONS = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff", ".dcm"}
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3
PATIENT_WAIT_TIMEOUT   = 900

# Expected field names in the active Access form
ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

# Name of the subform that lists documents
SFDOC_SUBFORM_NAME = "SFDoc"

# Description to use in the database for each file type; 
# default is "Image" except for TIFF which is "OCT" and DICOM which is "DICOM"
EXAM_DESCRIPTION = {
    ".jpg":  "Image",
    ".jpeg": "Image",
    ".jfif": "Image",
    ".png":  "Image",
    ".bmp":  "Image",
    ".tif":  "OCT",
    ".tiff": "OCT",
    ".dcm":  "DICOM",
}

# Logging is configured once at module level so that both the main thread and
# the worker thread share the same handlers.  The format includes the thread
# name, which makes it straightforward to distinguish producer events from
# consumer processing steps in the log output.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("image_router.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")

def db_connect(mdb_path: Path):
    """
    Open and return a pyodbc connection to a Microsoft Access database file.

    Uses the legacy JET/ACE ODBC driver that ships with Microsoft Office.
    Error handling is intentionally deferred to the caller so that each call
    site can log a context-specific message before propagating the failure.
    """
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> dict | None:
    """
    Interrogate the currently active Microsoft Access form via COM automation
    and return the patient's identifying fields if the expected form is open.

    The function attaches to the running Access process with GetActiveObject(),
    then iterates over every control on the foreground form looking for the
    three required fields (patient code, last name, first name).  Using index-
    based iteration rather than direct attribute access is necessary because the
    COM control collection does not support Python-style key lookup.

    Returns a dictionary with keys "code", "nom", and "prenom" when all three
    fields are present and readable, or None if:
      - win32com is not available (non-Windows platform),
      - no Access instance is running,
      - no form is currently in focus, or
      - the active form does not contain the expected patient fields.
    """
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in target:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                # Some controls raise COM errors on .Value access (e.g. labels);
                # silently skip them and continue scanning.
                pass

        if not target.issubset(data.keys()):
            return None

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }

    except Exception as e:
        log.debug(f"COM error: {e}")
        return None


def find_patient_folder(patient_code: str) -> Path | None:
    """
    Resolve the filesystem path of a patient's photo folder by querying
    the Documents table in PUBLIC.MDB.

    The 'Photo externe' column stores a relative path of the form
    '<group>\\<patient_folder>\\<filename>'.  Only the first two path
    components (group directory and patient sub-directory) are used to
    reconstruct the folder path under DEST_PHOTOS.

    Returns the resolved Path if the folder exists on disk, or None on any
    failure: missing pyodbc, missing database file, no matching record,
    malformed path value, or folder absent on disk.
    """
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available.")
        return None
    if not PUBLIC_MDB.exists():
        log.error(f"PUBLIC.MDB not found: {PUBLIC_MDB}")
        return None
    try:
        conn   = db_connect(PUBLIC_MDB)
        cursor = conn.cursor()
        # Fetch a single existing document to derive the folder path.
        # Selecting TOP 1 avoids pulling every document row for busy patients.
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            log.warning(f"No existing document found for patient {patient_code}.")
            return None

        # Strip leading/trailing backslashes before splitting so that the
        # resulting list starts with the group name rather than an empty string.
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error(f"Unexpected Photo externe format: {row[0]}")
            return None

        folder = DEST_PHOTOS / parts[0] / parts[1]
        if not folder.is_dir():
            log.error(f"Folder found in DB but missing on disk: {folder}")
            return None

        log.info(f"Patient folder resolved: {folder}")
        return folder
    except Exception as e:
        log.error(f"DB folder lookup failed: {e}")
        return None


def insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """
    Create a new row in the Documents table of PUBLIC.MDB to register the
    transferred image as an official document for the given patient.

    The inserted record mirrors the structure created by Studiovision itself:
      - TypeVW = 99 signals an external document (image linked by path).
      - TEXTE and Photo externe both store the relative path so that the
        application can both display the description and open the file.
      - NumDocExterne is left NULL because external numbering is not required.

    DOCUM.MDB is intentionally not used here; it is managed exclusively by
    Studiovision and is effectively read-only for external writers.

    Returns True if the commit succeeded, False on any error.
    """
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert skipped.")
        return False

    # IMPORTANT: target_mdb must be PUBLIC.MDB because DOCUM.MDB is read-only for this operation.
    target_mdb = PUBLIC_MDB
    if not target_mdb.exists():
        log.error("PUBLIC.MDB not found, insert skipped.")
        return False

    try:
        conn   = db_connect(target_mdb)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO Documents
                ([code patient], [Date], DESCRIPTIONS, TEXTE, [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, ?, 99, NULL)
            """,
            (int(patient["code"]), datetime.now(), description, relative_path, relative_path)
        )
        conn.commit()
        conn.close()
        log.info(f"Insert OK: patient={patient['code']} path='{relative_path}' db={target_mdb.name}")
        return True
    except Exception as e:
        log.error(f"DB insert failed: {e}")
        return False


# Numeric constant for the Access subform control type, used when iterating
# the Controls collection to identify embedded subform controls.
_AC_SUBFORM = 112


def _requery_form(form) -> None:
    """
    Recursively requery every subform in the control tree, then requery the
    parent form itself.

    Processing subforms before the parent ensures child record sources are
    refreshed before Access evaluates any link fields on the parent, avoiding
    stale-record mismatches.  Requery() forces Access to re-execute the
    underlying record source query; if it is unavailable (some runtime
    configurations block it), Refresh() is used as a graceful fallback to at
    least repaint the current data.
    """
    # Recurse into subforms first so their data is fresh before the parent is requeried.
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType == _AC_SUBFORM:
                _requery_form(ctrl.Form)
        except Exception:
            pass

    try:
        form.Requery()
        log.info(f"Requery() on '{form.Name}'")
    except Exception as e_req:
        log.warning(f"Requery() unavailable on '{form.Name}' ({e_req}), trying Refresh()...")
        try:
            form.Refresh()
            log.info(f"Refresh() on '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() also unavailable on '{form.Name}' ({e_ref})")


def _goto_last_record(form) -> None:
    """
    Navigate the SFDoc subform's recordset to its last row after a Requery.

    After Requery(), Access positions the cursor on the first record.  Calling
    MoveLast() ensures the clinician immediately sees the most recently added
    document rather than having to scroll down manually.  The function recurses
    into nested subforms to locate SFDoc regardless of how deeply it is nested
    in the form hierarchy.
    """
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue
            if ctrl.Name == SFDOC_SUBFORM_NAME:
                ctrl.Form.Recordset.MoveLast()
                log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
                return
            _goto_last_record(ctrl.Form)
        except Exception as e:
            log.debug(f"MoveLast failed on '{getattr(ctrl, 'Name', '?')}': {e}")


def refresh_ui() -> None:
    """
    Trigger a full requery of the active Access form and navigate to the last
    document record, making the newly inserted image visible to the clinician.

    The function is intentionally non-blocking: all COM exceptions are caught
    and logged as warnings so that a transient Access error never causes the
    worker thread to abort file processing.
    """
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return
        _requery_form(form)
        _goto_last_record(form)
    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


def wait_for_file(file: Path) -> bool:
    """
    Poll until the file can be opened for binary reading, or until the
    maximum number of attempts is exhausted.

    Medical imaging devices often write files incrementally, leaving them
    locked by the writing process for several seconds after the filesystem
    creation event fires.  Attempting to read the file immediately would
    result in partial data or a PermissionError.  This function retries up to
    FILE_LOCK_MAX_ATTEMPTS times with FILE_LOCK_RETRY_DELAY seconds between
    each attempt, providing a safe stabilisation window.

    Returns True as soon as the file is readable, or False if it remains
    locked after all attempts, signalling the caller to abandon processing.
    """
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retrying...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error(f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {file}")
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> Path | None:
    """
    Atomically move a file to dest_folder, creating the destination directory
    tree if it does not yet exist.

    When a file with the same name already exists at the destination (e.g. two
    scans taken within the same second), a Unix timestamp suffix is appended to
    the stem to guarantee a unique filename and avoid silent data loss.

    The optional label parameter is prepended to the log entry in square
    brackets, making it easy to filter orphaned files in the log.

    Returns the final destination Path on success, or None if shutil.move()
    raises an exception (e.g. cross-device rename that also fails to copy).
    """
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name

    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict, renamed to {dest.name}")

    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} -> {dest}")
        return dest
    except Exception as e:
        log.error(f"Move failed: {e}")
        return None


def orphan_file(file: Path) -> None:
    """
    Quarantine a file that could not be associated with a patient by moving it
    to ORPHAN_DIR.

    Orphaning is preferred over deletion because it preserves the original data
    for manual review.  The calling context (e.g. patient lookup timeout, folder
    resolution failure) is logged by the caller before this function is invoked.
    """
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    """
    Long-running consumer that drains the shared file queue and processes each
    image through the complete pipeline.

    COM must be initialised per-thread with CoInitialize() before any win32com
    calls; CoUninitialize() in the finally block releases the apartment
    regardless of how the thread exits.

    Processing pipeline for each file:
      1. Existence check — the file may have been deleted between the Watchdog
         event and the worker picking it from the queue.
      2. Lock-wait — polls until the file is fully written (see wait_for_file).
      3. Patient identification — polls the active Access form until a patient
         record is open or the timeout expires.  Files whose timeout elapses
         are orphaned to avoid stalling the queue indefinitely.
      4. Folder resolution — queries PUBLIC.MDB to find the patient's directory.
      5. File move — transfers the image to the patient folder.
      6. DB insert — registers the image in the Documents table.
      7. UI refresh — triggers a 1.5-second-delayed Requery so Access picks up
         the committed row before the refresh call is made.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    try:
        while True:
            try:
                file: Path = file_queue.get()
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue

            log.info(f"Processing: {file.name} ({file_queue.qsize()} pending)")

            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                file_queue.task_done()
                continue

            # Patient identification loop: poll the Access form at fixed intervals
            # until a patient is found or the configurable timeout is reached.
            # The first_log flag prevents flooding the log on every poll cycle.
            patient    = None
            start_time = time.monotonic()
            first_log  = True

            while True:
                patient = get_active_patient()
                if patient:
                    break

                elapsed = time.monotonic() - start_time
                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    orphan_file(file)
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(f"No patient open, waiting (timeout in {PATIENT_WAIT_TIMEOUT // 60} min)")
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(f"Patient: {patient['nom']} {patient['prenom']} (code {patient['code']})")

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(f"Could not resolve folder for patient {patient['code']}. Orphaning.")
                orphan_file(file)
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
                # Give Access enough time to commit the record and make it
                # visible to Requery() before the UI refresh call is made.
                time.sleep(1.5)
                refresh_ui()
            else:
                log.warning("Insert failed, refresh skipped.")

            file_queue.task_done()

    finally:
        pythoncom.CoUninitialize()

class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler that acts as the producer side of the pipeline.

    Each time a new file is detected under SOURCE_DIR, on_created() checks
    whether its extension is in the watched set and, if so, pushes the Path
    onto the shared queue for the worker thread to consume.  Directory creation
    events are explicitly ignored because they carry no image data.
    """
    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info(f"Enqueued: {file.name} (queue size: {self._queue.qsize() + 1})")
        self._queue.put(file)


def main() -> None:
    """
    Application entry point: validate configuration, start the worker thread
    and Watchdog observer, then block until interrupted.

    On KeyboardInterrupt (Ctrl+C) the observer is stopped cleanly and the
    worker is given a chance to finish any files still in the queue before
    the process exits, preventing mid-transfer data loss.
    """
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 1 started")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Dest       : {DEST_PHOTOS}")
    log.info(f"  PUBLIC.MDB : {PUBLIC_MDB}")
    log.info(f"  DOCUM.MDB  : {DOCUM_MDB}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext        : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(target=worker, args=(file_queue,), name="Worker", daemon=True)
    worker_thread.start()

    producer = ImageProducer(file_queue)
    observer = Observer()
    observer.schedule(producer, str(SOURCE_DIR), recursive=True)
    observer.start()
    log.info("Watching for images. Press Ctrl+C to stop.")

    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Shutdown requested.")
    finally:
        observer.stop()
        observer.join()

        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Waiting for {remaining} remaining file(s)...")
            file_queue.join()

        log.info("Image Router stopped.")

if __name__ == "__main__":
    main()