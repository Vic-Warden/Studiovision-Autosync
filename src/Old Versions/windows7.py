"""
windows7.py
Image Router — Windows 7 compatible build.

Functionally equivalent to studiovision_monitor.py but written to remain
compatible with Python 3.9 on Windows 7 workstations, where the union-type
hint syntax (X | Y) is not available.  All type annotations therefore use
typing.Optional rather than the native union operator, and f-strings are used
only where the minimum Python version guarantees support.

The overall pipeline is identical to the main monitor build:
  - A Watchdog observer detects new image files in SOURCE_DIR.
  - Each detected file is placed on a thread-safe queue.
  - A single worker thread processes files sequentially through the
    lock-wait → patient lookup → move → DB insert → UI refresh pipeline.
  - Files that cannot be matched to an open patient within the configurable
    timeout are quarantined in ORPHAN_DIR.

See studiovision_monitor.py for detailed architecture notes.
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
from typing import Optional
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

# Logging is configured once at module level and shared across all threads.
# The thread name is included in the format string to distinguish producer
# (main thread) events from consumer (Worker thread) processing steps.
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

# Helper to connect to an Access MDB with pyodbc, with error handling deferred to caller
def db_connect(mdb_path: Path):
    """
    Open and return a pyodbc connection to a Microsoft Access database file.

    Uses the JET/ACE ODBC driver bundled with Microsoft Office.  Error
    handling is deferred to the caller so that each call site can log a
    descriptive message before propagating the failure.
    """
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> Optional[dict]:
    """
    Interrogate the currently active Microsoft Access form via COM automation
    and return the patient's identifying fields if the expected form is open.

    Attaches to a running Access instance with GetActiveObject() and iterates
    the foreground form's Controls collection, collecting the values of the
    three mandatory patient fields.  Index-based iteration is required because
    the COM Controls object does not support Python-style key access.

    Returns a dict with keys "code", "nom", and "prenom" on success, or None
    when win32com is unavailable, no Access instance is running, no form is
    active, or the active form does not contain all required patient fields.
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
                # Label controls and non-data controls raise COM errors on
                # .Value access; skip them silently.
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


def find_patient_folder(patient_code: str) -> Optional[Path]:
    """
    Resolve the filesystem path of a patient's photo folder by querying
    PUBLIC.MDB.

    The 'Photo externe' column stores a relative UNC-style path of the form
    '<group>\\<patient>\\<filename>'.  Only the first two path components are
    used to reconstruct the folder path under DEST_PHOTOS.

    Returns the resolved Path if the folder exists on disk, or None on any
    failure: missing pyodbc, missing database, no matching record, malformed
    path value, or directory not found on disk.
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
        # TOP 1 avoids retrieving all documents for patients with many records.
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

        # Strip enclosing backslashes before splitting to avoid empty leading
        # elements in the resulting list.
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
    Create a new row in the Documents table of PUBLIC.MDB to register a
    transferred image as an official Studiovision document.

    TypeVW = 99 denotes an externally linked image.  Both TEXTE and
    'Photo externe' carry the relative path so that Studiovision can display
    the description and open the file from the same stored value.
    NumDocExterne is NULL because external sequential numbering is not used.

    DOCUM.MDB is not written to; it is exclusively managed by Studiovision.

    Returns True on successful commit, False on any error.
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


# Numeric constant for the Access subform control type, used when traversing
# the Controls collection to distinguish embedded subforms from other controls.
_AC_SUBFORM = 112


def _requery_form(form) -> None:
    """
    Recursively requery every subform before requerying the parent form.

    Refreshing child recordsets before the parent prevents stale link-field
    values from being used when Access re-evaluates master/child relationships.
    Falls back to Refresh() if Requery() is not available in the current
    runtime configuration.
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
    Navigate the SFDoc subform's recordset to its last row.

    After a Requery(), Access positions the cursor on record #1.  This
    function locates SFDoc by recursing through the control tree and calls
    MoveLast() so the clinician immediately sees the newly added document.
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
    document row so the newly inserted image is immediately visible.

    Deliberately non-blocking: all COM exceptions are caught and logged as
    warnings so that a transient Access error cannot stall the worker thread.
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
    Poll until the file can be opened for binary reading or the maximum
    number of retry attempts is exhausted.

    Imaging devices often keep a file locked for several seconds after the
    filesystem creation event fires.  Opening too early produces partial data
    or a PermissionError.  This function retries up to FILE_LOCK_MAX_ATTEMPTS
    times with FILE_LOCK_RETRY_DELAY seconds between each attempt.

    Returns True when the file is readable, or False if it remains locked
    throughout all attempts.
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


def move_file(source: Path, dest_folder: Path, label: str = "") -> Optional[Path]:
    """
    Atomically move source to dest_folder, creating intermediate directories
    as needed and resolving filename conflicts with a Unix timestamp suffix.

    A conflict arises when two images for the same patient arrive within a
    single second.  Appending the timestamp guarantees uniqueness without
    discarding either file.  The optional label is bracketed in the log entry
    for easy filtering (e.g. "[ORPHAN]").

    Returns the final destination Path on success, or None on failure.
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
    Quarantine a file that could not be matched to a patient by moving it to
    ORPHAN_DIR rather than deleting it.

    Quarantining preserves the original image data for manual recovery and
    audit, which is essential in a medical imaging context.
    """
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    """
    Long-running consumer thread that drains the shared file queue and
    processes each image through the complete pipeline.

    COM must be initialised per-thread; CoInitialize() is called at entry and
    CoUninitialize() is guaranteed via the finally block regardless of how the
    thread exits.

    Processing pipeline per file:
      1. Existence check — the file may be deleted between the Watchdog event
         and the moment the worker dequeues it.
      2. Lock-wait — polls until the file is fully written and readable.
      3. Patient loop — polls the active Access form until a patient is found
         or PATIENT_WAIT_TIMEOUT elapses; orphans the file on timeout.
      4. Folder resolution — queries PUBLIC.MDB for the patient's directory.
      5. File move — transfers the image into the resolved patient folder.
      6. DB insert — registers the image in the Documents table.
      7. UI refresh — waits 1.5 s for Access to see the committed record,
         then calls refresh_ui() to update the form.
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

            # Patient identification loop: poll the foreground Access form at
            # PATIENT_POLL_INTERVAL second intervals.  The first_log flag
            # suppresses duplicate "waiting" messages on subsequent poll cycles.
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
                # A short delay gives Access time to make the committed row
                # visible before Requery() is called on the form.
                time.sleep(1.5)
                refresh_ui()
            else:
                log.warning("Insert failed, refresh skipped.")

            file_queue.task_done()

    finally:
        pythoncom.CoUninitialize()

class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler — producer side of the producer/consumer pipeline.

    Receives filesystem creation events from the Observer and filters them to
    image files of supported types before pushing their paths onto the shared
    queue for the worker thread to consume.  Directory creation events are
    explicitly ignored because they carry no image payload.
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
    Application entry point: validate runtime prerequisites, start the worker
    thread and Watchdog observer, then block until a keyboard interrupt.

    On shutdown the observer is stopped first to prevent new files from being
    enqueued while the worker is draining.  If the queue is non-empty at that
    point, the main thread waits for it to drain completely before exiting, so
    no in-flight file is abandoned mid-pipeline.
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