"""
box1V3.py
Image Router — Version 3, Box 1 (standard imaging device build).

This variant targets the first imaging station ("Box 1") and adds all Version 3
improvements on top of the baseline monitor:

  1. PollingObserver — reliable detection on network/UNC shares.
  2. Network share wait at startup — blocks until SOURCE_DIR is reachable.
  3. Auto-reconnect main loop — restarts the observer after network drops.
  4. Targeted parent Refresh() + SFDoc-only Requery() — avoids resetting the
     Access record pointer to #1 when the document list is refreshed.
  5. Dirty-state guard — clears form.Dirty before Requery() so Access does not
     silently ignore the call while a record is in edit mode.
  6. Requery retry loop — retries up to 3 times (0.5 s apart) before falling
     back to Refresh().
  7. Patient-code guard — skips the UI refresh if the operator navigated to a
     different patient during the burst debounce window.
  8. Log written to ~/studiovision/ — valid both as a script and as a .exe.

Architecture — producer/consumer with burst debounce:
  A PollingObserver enqueues new image Paths; a single worker thread processes
  them and defers UI refresh until the queue has been idle for 1.5 seconds so
  that a rapid burst of files triggers only one Access Requery.

Dependencies: watchdog, pyodbc, pywin32, pythoncom
"""

import os
import pythoncom
import queue
import shutil
import sys
import threading
import time
import ctypes
import logging
from datetime import datetime
from pathlib import Path
from watchdog.observers.polling import PollingObserver as Observer  # PollingObserver works on network shares; the default observer misses events on UNC/mapped paths
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

# Description inserted into the database per file type
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

# Log file is written to ~/studiovision/ so the path is valid both when running
# as a plain script and when packaged as a compiled .exe via PyInstaller.
_LOG_DIR  = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, "image_router.log")

# Logging is configured once at module level and shared across all threads.
# Including the thread name in the format string makes it easy to separate
# producer (main thread) events from consumer (Worker thread) activity.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")


def db_connect(mdb_path: Path):
    """
    Open and return a pyodbc connection to a Microsoft Access database file.

    Uses the JET/ACE ODBC driver bundled with Microsoft Office. Error handling
    is intentionally deferred to the caller so each call site can provide a
    context-specific error message before propagating the failure.
    """
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> dict | None:
    """
    Interrogate the currently active Microsoft Access form via COM automation
    and return the patient's identifying fields if the expected form is open.

    Attaches to a running Access instance with GetActiveObject() and iterates
    the foreground form's Controls collection to locate the three mandatory
    patient fields (code, last name, first name).  Index-based iteration is
    required because the COM Controls object does not support key access.

    Returns a dict with keys "code", "nom", and "prenom" on success, or None
    when win32com is unavailable, no Access instance is running, no form is
    active, or the active form lacks the expected patient fields.
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
                # Label controls and decorative controls raise COM errors on
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


def find_patient_folder(patient_code: str) -> Path | None:
    """
    Resolve the filesystem path of a patient's photo folder by querying
    the Documents table in PUBLIC.MDB.

    The 'Photo externe' column stores a relative path of the form
    '<group>\\<patient_folder>\\<filename>'.  Only the first two components
    are used to reconstruct the folder under DEST_PHOTOS.

    Returns the resolved Path if the folder exists on disk, or None on any
    failure: missing pyodbc, missing database, no matching record, malformed
    path value, or directory absent from disk.
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

        # Strip leading/trailing backslashes before splitting so the list
        # starts with the group name rather than an empty first element.
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
    Create a new row in PUBLIC.MDB's Documents table to register the
    transferred image as an official Studiovision document.

    TypeVW = 99 signals an externally linked image.  Both TEXTE and
    'Photo externe' store the relative path so Studiovision can display the
    description and open the file.  NumDocExterne is NULL because external
    sequential numbering is not required.

    DOCUM.MDB is intentionally not written; it is managed exclusively by
    Studiovision and is effectively read-only for external processes.

    Returns True on successful commit, False on any error.
    """
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert skipped.")
        return False

    if not PUBLIC_MDB.exists():
        log.error("PUBLIC.MDB not found, insert skipped.")
        return False

    try:
        conn   = db_connect(PUBLIC_MDB)
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
        log.info(f"Insert OK: patient={patient['code']} path='{relative_path}' db={PUBLIC_MDB.name}")
        return True
    except Exception as e:
        log.error(f"DB insert failed: {e}")
        return False


# Numeric constant for the Access subform control type, used when traversing
# the Controls collection to identify embedded subforms.
_AC_SUBFORM = 112


def _find_sfdoc(form):
    """
    Recursively search the form's control tree for the subform named
    SFDOC_SUBFORM_NAME and return its inner Form object, or None if not found.

    Targeting SFDoc directly — rather than requerying the entire parent form
    — avoids resetting the parent's record pointer to record #1, which would
    disrupt the clinician's current navigation position.  The recursion handles
    arbitrarily deep nesting without making assumptions about the form layout.
    """
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue
            if ctrl.Name == SFDOC_SUBFORM_NAME:
                return ctrl.Form
            found = _find_sfdoc(ctrl.Form)
            if found is not None:
                return found
        except Exception:
            # COM errors on ControlType or Name access are silently skipped;
            # the search continues to the next control in the collection.
            pass
    return None


def refresh_ui(expected_patient_code: str | None = None) -> None:
    """
    Two-step refresh strategy:
      1. form.Refresh() on the parent — repaints bound image controls without
         moving the record pointer (change 4).
      2. Dirty guard — clears edit mode before Requery() since Access silently
         ignores Requery() while a record is being edited (change 5).
      3. Requery() on SFDoc with 3-attempt retry (0.5 s apart); falls back to
         sfdoc.Refresh() if all attempts fail (change 6).
      4. MoveLast() to position on the newly added document row.

    If expected_patient_code is provided, verifies the correct patient is still
    on screen before acting — guards against user switching records during the
    1.5 s debounce window (change 7).

    All COM errors are caught and logged as non-blocking warnings.
    """
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        # Change 7: patient code guard
        if expected_patient_code is not None:
            current = get_active_patient()
            if current is None or current["code"] != expected_patient_code:
                current_code = current["code"] if current else "none"
                log.warning(
                    f"Refresh skipped: expected patient {expected_patient_code} "
                    f"but current patient is {current_code}."
                )
                return

        # Change 4: refresh parent form to repaint image controls
        try:
            form.Refresh()
            log.info(f"Refresh() on parent form '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() on parent form failed ({e_ref}), continuing...")

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found in the active form. "
                "SFDoc refresh skipped."
            )
            return

        # Change 5: dirty state guard
        try:
            if form.Dirty:
                log.info("Parent form is dirty — clearing edit mode before Requery().")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Could not read/clear form.Dirty: {e_dirty}")

        # Change 6: Requery retry loop
        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5
        requeried = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}' (attempt {attempt}/{_REQUERY_ATTEMPTS})")
                requeried = True
                break
            except Exception as e_req:
                log.warning(
                    f"Requery() attempt {attempt}/{_REQUERY_ATTEMPTS} failed "
                    f"on '{SFDOC_SUBFORM_NAME}': {e_req}"
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requeried:
            log.warning(
                f"All Requery() attempts failed on '{SFDOC_SUBFORM_NAME}', "
                "falling back to Refresh()."
            )
            try:
                sfdoc.Refresh()
                log.info(f"Fallback Refresh() on '{SFDOC_SUBFORM_NAME}'")
            except Exception as e_ref2:
                log.warning(
                    f"Fallback Refresh() also failed on '{SFDOC_SUBFORM_NAME}': {e_ref2}"
                )

        # Navigate to the last record to make the new document visible
        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


def wait_for_file(file: Path) -> bool:
    """Poll until the file is readable, retrying up to FILE_LOCK_MAX_ATTEMPTS times.

    Medical imaging devices hold a write lock on the file for several seconds
    after the filesystem creation event fires.  Attempting to read too early
    yields partial data or a PermissionError.  This function provides the
    stabilisation window by retrying at FILE_LOCK_RETRY_DELAY second intervals.

    Returns True as soon as the file opens successfully, or False if it remains
    locked after all allowed attempts.
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
    """Move source to dest_folder, appending a Unix timestamp on filename conflicts.

    Creating the destination directory tree with mkdir(parents=True) ensures no
    manual setup is required for new patient folders.  Timestamp suffixing
    prevents silent data loss when two images arrive for the same patient within
    one second.  The optional label is bracketed in the log line for easy
    filtering (e.g. "[ORPHAN]").

    Returns the final destination Path on success, or None if the move fails.
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
    """Quarantine a file that could not be matched to a patient.

    Moving to ORPHAN_DIR rather than deleting preserves the original image
    data for manual review and audit, which is critical in a medical context.
    """
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    """
    Long-running consumer thread that drains the shared queue and processes
    each image through the complete pipeline.

    COM must be initialised per-thread; CoInitialize() is called on entry and
    CoUninitialize() is guaranteed via the finally block regardless of how the
    thread exits.

    Burst debounce: the worker uses a 1.5-second blocking get() timeout.  While
    files arrive faster than that interval, it processes them back-to-back with
    needs_refresh=True.  Once the queue is idle for 1.5 seconds (burst end), a
    single UI refresh is fired for the last successfully inserted patient, rather
    than calling Access COM after every individual insert.

    Processing pipeline per file:
      1. Existence check — the file may be deleted between the Watchdog event
         and the moment it is dequeued.
      2. Lock-wait — polls until the file is fully written and readable.
      3. Patient identification loop — polls the active Access form until a
         patient is found or PATIENT_WAIT_TIMEOUT elapses; orphans on timeout.
      4. Folder resolution — queries PUBLIC.MDB for the patient's directory.
      5. File move — transfers the image to the resolved patient folder.
      6. DB insert — registers the image in the Documents table.
      7. Burst refresh flag — sets needs_refresh=True and records the patient
         code for the deferred UI update at burst end.

    On worker shutdown any pending refresh is flushed in the finally block to
    prevent the last batch of inserts from going unrefreshed.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh: bool = False
    last_patient_code: str | None = None

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                # Queue has been idle for 1.5 s — the burst is over.
                # Fire the deferred UI refresh now so Access is updated once
                # rather than after every individual insert.
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh = False
                    last_patient_code = None
                continue
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
                    log.info(
                        f"No patient open, waiting "
                        f"(timeout in {PATIENT_WAIT_TIMEOUT // 60} min)"
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                f"Patient: {patient['nom']} {patient['prenom']} "
                f"(code {patient['code']})"
            )

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(
                    f"Could not resolve folder for patient {patient['code']}. "
                    "Orphaning."
                )
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
                needs_refresh     = True
                last_patient_code = patient["code"]
                log.debug(
                    f"Insert OK — needs_refresh=True, "
                    f"last_patient_code={last_patient_code} "
                    "(refresh deferred to burst end)."
                )
            else:
                log.warning("Insert failed, refresh flag unchanged.")

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler — producer side of the producer/consumer pipeline.

    Receives filesystem creation events from the PollingObserver, filters them
    to supported image extensions, and enqueues the Path on the shared queue for
    the worker thread to process.  Directory events are ignored because they
    carry no image data.
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


def wait_for_network_share() -> None:
    """
    Block at startup until SOURCE_DIR is accessible.

    Returns immediately for local drive paths (identified by a colon at index 1,
    e.g. C:\\...).  For network/UNC paths the function polls is_dir() every 10
    seconds, logging a warning on each failed attempt.  Once accessible, it logs
    an info message confirming recovery so the log reflects how long the share
    was unavailable.
    """
    source_str = str(SOURCE_DIR)
    is_unc   = source_str.startswith("\\\\") or source_str.startswith("//")
    is_local = not is_unc and len(source_str) >= 2 and source_str[1] == ":"

    if is_local:
        return

    first_attempt = True
    while True:
        try:
            if SOURCE_DIR.is_dir():
                if not first_attempt:
                    log.info(f"Network share is now reachable: {SOURCE_DIR}")
                return
        except Exception:
            pass
        log.warning(f"Network share not reachable, retrying in 10 s: {SOURCE_DIR}")
        first_attempt = False
        time.sleep(10)


def prevent_sleep() -> None:
    """Prevent Windows from sleeping or turning off the display while the program runs."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning(f"Could not set execution state: {e}")


def main() -> None:
    """
    Application entry point for Box 1, Version 3.

    Startup sequence:
      1. Block until SOURCE_DIR is reachable (network share wait).
      2. Create ORPHAN_DIR and log startup configuration.
      3. Start the worker (consumer) thread.
      4. Start the PollingObserver (producer) on SOURCE_DIR.
      5. Enter the auto-reconnect main loop: detects a dead observer thread
         (e.g. after a network drop), waits for the share to recover, and
         restarts a fresh observer automatically.

    On KeyboardInterrupt (Ctrl+C) the observer is stopped first to halt new
    enqueues, then the main thread waits for the queue to drain completely so
    no in-flight file is abandoned mid-pipeline.
    """
    prevent_sleep()
    # Network share check before doing anything else.
    log.info("Checking network share availability...")
    wait_for_network_share()

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 3 started")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Dest       : {DEST_PHOTOS}")
    log.info(f"  PUBLIC.MDB : {PUBLIC_MDB}")
    log.info(f"  DOCUM.MDB  : {DOCUM_MDB}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext        : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    )
    worker_thread.start()

    # Auto-reconnect loop (improvement 3): seconds to wait before restarting
    # the observer after a network drop.
    _RECONNECT_DELAY = 15

    def _start_observer() -> Observer:
        """Create, schedule, and start a fresh PollingObserver on SOURCE_DIR."""
        producer = ImageProducer(file_queue)
        obs = Observer()
        obs.schedule(producer, str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started. Watching for images. Press Ctrl+C to stop.")
        return obs

    observer = _start_observer()

    try:
        while True:
            time.sleep(1)
            if not observer.is_alive():
                log.warning(
                    f"Observer died (possible network drop). "
                    f"Waiting {_RECONNECT_DELAY} s before reconnecting..."
                )
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                time.sleep(_RECONNECT_DELAY)
                log.info("Restarting Observer after network recovery.")
                observer = _start_observer()

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