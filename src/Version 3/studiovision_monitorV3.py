"""
studiovision_monitorV3.py
Image Router — Version 3, latest feature build for modern Windows.

Extends the baseline monitor with the following improvements over earlier versions:

  1. PollingObserver — replaces the default inotify/ReadDirectoryChanges observer
     with a polling-based one that works reliably on network shares (UNC paths
     and mapped drives), which the standard observer misses.

  2. Network share wait — blocks at startup until SOURCE_DIR is reachable, so the
     process does not crash immediately if the file server is temporarily offline.

  3. Auto-reconnect loop — if the Watchdog observer thread dies (e.g. because the
     network share goes offline mid-session), the main loop detects it, waits for
     the share to come back, and restarts a fresh observer automatically.

  4. Targeted parent Refresh() — instead of requerying the entire Access form
     hierarchy (which resets the record pointer to #1), only the parent form
     receives a Refresh() (repaint only) while SFDoc receives Requery() +
     MoveLast() so the document list is reloaded without disrupting navigation.

  5. Dirty-state guard — clears form.Dirty before calling Requery() to prevent
     Access from silently ignoring the refresh while a record is in edit mode.

  6. Requery retry loop — retries Requery() up to three times with a 0.5-second
     delay between attempts before falling back to Refresh().

  7. Patient-code guard — compares the patient code captured at insert time with
     the one currently on screen before refreshing, so the UI is not disturbed if
     the operator navigated to a different patient during the 1.5-second debounce.

  8. Log in ~/studiovision/ — writes the log file to the user's home directory so
     that the path is valid both when running as a script and as a compiled .exe.

Architecture — producer/consumer, single worker thread:
  Producer: Watchdog PollingObserver on SOURCE_DIR (recursive).
  Consumer: worker() thread draining a thread-safe queue with burst debounce
            (refreshes Access once after 1.5 s of queue idle time rather than
            after every individual insert).

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
from watchdog.observers.polling import PollingObserver as Observer
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
WATCHED_EXTENSIONS = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff", ".dcm", ".pdf", ".rtf", ".doc", ".docx", ".odt"}
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

    ".pdf":  "Document",
    ".rtf":  "Document",
    ".doc":  "Document",
    ".docx": "Document",
    ".odt":  "Document"
}

# Configure logging to file and console with timestamps and thread names
_LOG_DIR  = Path(os.path.expanduser("~")) / "studiovision"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / "image_router.log"

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


_NETWORK_SHARE_POLL = 10  # seconds between retries when share is unreachable


def wait_for_network_share() -> None:
    """
    Block until SOURCE_DIR is accessible.

    For local paths the check passes immediately.  For UNC/network shares the
    function loops every _NETWORK_SHARE_POLL seconds, logging a warning on
    each failed attempt, so the programme waits silently until the share comes
    online rather than crashing at startup.
    """
    is_network = str(SOURCE_DIR).startswith("\\\\") or str(SOURCE_DIR).startswith("//")

    if not is_network:
        return  # nothing to wait for on local paths

    attempt = 0
    while not SOURCE_DIR.is_dir():
        attempt += 1
        log.warning(
            f"Network share not reachable: {SOURCE_DIR}  "
            f"(attempt {attempt}, retrying in {_NETWORK_SHARE_POLL}s)"
        )
        time.sleep(_NETWORK_SHARE_POLL)

    if attempt:
        log.info(f"Network share is now accessible after {attempt} attempt(s): {SOURCE_DIR}")


def db_connect(mdb_path: Path):
    """Helper to connect to an Access MDB with pyodbc."""
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> dict | None:
    """
    Returns a dict with patient info if an Access form with the expected fields
    is active, else None.
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
    Uses PUBLIC.MDB Documents table to resolve the patient's photo folder.
    Returns a Path if successful or None if any step fails.
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
    Inserts a new record into PUBLIC.MDB Documents for the given patient,
    relative path, and description.
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


# Access constant for subform control type
_AC_SUBFORM = 112


def _find_sfdoc(form):
    """
    Recursively walks the form's control tree and returns the Form object
    of the subform named SFDOC_SUBFORM_NAME, or None if not found.
    This avoids touching the parent form's Recordset (and its current-record
    pointer), which prevents the "sent back to record #1" regression.
    """
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue
            if ctrl.Name == SFDOC_SUBFORM_NAME:
                return ctrl.Form
            # Descend into nested subforms
            found = _find_sfdoc(ctrl.Form)
            if found is not None:
                return found
        except Exception:
            pass
    return None


def refresh_ui(expected_patient_code: str | None = None) -> None:
    """
    Two-step refresh strategy:
      1. Refresh() on the PARENT form — updates bound image controls (photo
         display) without resetting its current-record pointer. We use
         Refresh() instead of Requery() on the parent to avoid jumping back
         to record #1.
      2. Requery() + MoveLast() on SFDoc only — reloads the document list
         and positions it on the newly added entry.

    Guards applied before refreshing:
      - Patient code guard  : if expected_patient_code is given, the current
        patient is re-checked; a mismatch skips the refresh entirely.
      - Dirty state guard   : Access silently ignores Requery() while a record
        is being edited. form.Dirty is cleared first to exit edit mode.
      - Requery retry loop  : 3 attempts with 0.5 s delay; falls back to
        Refresh() if all attempts fail.

    All COM errors are caught and logged so the worker thread is never blocked.
    """
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        # Guard: if the active patient has changed since the insert, skip the
        # refresh to avoid updating the wrong consultation on screen.
        if expected_patient_code is not None:
            current = get_active_patient()
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    f"Refresh skipped: active patient changed "
                    f"(expected={expected_patient_code}, current={current_code})."
                )
                return

        # Refresh() repaints bound controls from the current record without
        # moving the record pointer, so the consultation stays in place.
        try:
            form.Refresh()
            log.info(f"Refresh() on parent form '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() on parent form failed ({e_ref}), continuing...")

        # Requery SFDoc to load the new document row, then navigate to the last record.
        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found in the active form. "
                "SFDoc refresh skipped."
            )
            return

        # If the parent form has unsaved edits (Dirty=True), clear the dirty state
        # before calling Requery to prevent Access from raising a save-prompt dialog.
        try:
            if form.Dirty:
                log.info("Parent form is in edit mode (Dirty=True); clearing Dirty before Requery.")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Dirty check/clear failed ({e_dirty}), continuing...")

        # Requery SFDoc with up to _REQUERY_ATTEMPTS retries in case the COM
        # call fails transiently; fall back to Refresh() if all attempts fail.
        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5  # seconds

        requery_ok = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}' (attempt {attempt})")
                requery_ok = True
                break
            except Exception as e_req:
                log.warning(
                    f"Requery() attempt {attempt}/{_REQUERY_ATTEMPTS} failed "
                    f"on '{SFDOC_SUBFORM_NAME}': {e_req}"
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requery_ok:
            log.warning(
                f"All {_REQUERY_ATTEMPTS} Requery() attempts failed on "
                f"'{SFDOC_SUBFORM_NAME}'; falling back to Refresh()."
            )
            try:
                sfdoc.Refresh()
                log.info(f"Fallback Refresh() on '{SFDOC_SUBFORM_NAME}'")
            except Exception as e_ref2:
                log.warning(
                    f"Fallback Refresh() also failed on '{SFDOC_SUBFORM_NAME}': {e_ref2}"
                )

        # Navigate to the last record so the new document is visible.
        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


def wait_for_file(file: Path) -> bool:
    """Tries to open the file for reading to check if it's still locked."""
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
    """Moves the file to dest_folder, resolving name conflicts with a timestamp."""
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
    """Moves the file to the orphan folder with a warning log."""
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue: queue.Queue) -> None:
    """
    Processes files from the queue. Runs the full pipeline (lock-wait →
    patient lookup → move → DB insert) for each file, then fires a single
    UI refresh once the queue has been idle for 1.5 s (burst debounce).
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
                # Queue drained — if at least one insert succeeded since the
                # last refresh, now is the right moment to update the UI.
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
                needs_refresh = True
                last_patient_code = patient["code"]
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")

            file_queue.task_done()

    finally:
        # Flush any pending refresh before the thread exits
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler — producer side of the producer/consumer pipeline.

    Receives filesystem creation events from the PollingObserver and filters them
    to supported image extensions before enqueuing the path for the worker thread.
    Directory creation events are ignored since they carry no image data.
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
    Application entry point for Version 3.

    Startup sequence:
      1. Wait for the network share (improvement 2) so the process does not
         crash if the file server is temporarily offline at launch.
      2. Validate SOURCE_DIR and create ORPHAN_DIR if needed.
      3. Start the worker thread and the Watchdog PollingObserver.
      4. Enter the auto-reconnect main loop (improvement 3): if the observer
         thread dies (e.g. network drop), the loop stops the dead observer,
         waits for the share to come back, and restarts a fresh observer.

    On KeyboardInterrupt (Ctrl+C) the observer is stopped cleanly and the
    main thread waits for any queued files to finish processing before exit.
    """

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 3 started")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Dest       : {DEST_PHOTOS}")
    log.info(f"  PUBLIC.MDB : {PUBLIC_MDB}")
    log.info(f"  DOCUM.MDB  : {DOCUM_MDB}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    log.info(f"  Log file   : {_LOG_FILE}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext        : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    )
    worker_thread.start()

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — watching for images. Press Ctrl+C to stop.")
        return obs

    observer = _start_observer()

    # Auto-reconnect loop: if the Watchdog observer dies (e.g. due to a
    # temporary network drop on a UNC share), wait and restart it automatically.
    _RECONNECT_WAIT = 15  # seconds to pause before restarting the observer

    try:
        while True:
            if not observer.is_alive():
                log.warning("Observer has stopped (network drop?). Attempting reconnect...")
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass

                # Re-check the share before spawning a new observer (improvement 2).
                wait_for_network_share()
                log.info(f"Waiting {_RECONNECT_WAIT}s before restarting observer...")
                time.sleep(_RECONNECT_WAIT)

                observer = _start_observer()

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