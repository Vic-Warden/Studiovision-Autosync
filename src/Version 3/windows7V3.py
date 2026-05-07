"""
windows7V3.py
Image Router — Version 3, Windows 7 compatible build.

Combines all improvements introduced in Version 3 (PollingObserver, network
share wait, auto-reconnect, targeted SFDoc Requery, dirty-state guard, Requery
retry loop, patient-code guard, log in ~/studiovision/) with the Windows 7
Python 3.9 compatibility constraints of the windows7.py baseline (typing.Optional
instead of X | Y union syntax, str.format() where needed, daemon=True set as an
attribute rather than a constructor keyword).

See studiovision_monitorV3.py for a detailed description of each improvement.
"""

import os
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
from watchdog.observers.polling import PollingObserver as Observer   # change 1: PollingObserver for network shares
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

# Change 8: log file in ~/studiovision/ so it works correctly as a compiled executable
_LOG_DIR  = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, "image_router.log")

# Configure logging to file and console with timestamps and thread names
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


def db_connect(mdb_path):
    """Helper to connect to an Access MDB with pyodbc."""
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(mdb_path) + ";"
    )


def get_active_patient():
    # type: () -> Optional[dict]
    """Returns patient info from the active Access form, or None if unavailable."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data   = {}

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
        log.debug("COM error: %s", e)
        return None


def find_patient_folder(patient_code):
    # type: (str) -> Optional[Path]
    """Resolves the patient's photo folder from PUBLIC.MDB. Returns a Path or None."""
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available.")
        return None
    if not PUBLIC_MDB.exists():
        log.error("PUBLIC.MDB not found: %s", PUBLIC_MDB)
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
            log.warning("No existing document found for patient %s.", patient_code)
            return None

        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error("Unexpected Photo externe format: %s", row[0])
            return None

        folder = DEST_PHOTOS / parts[0] / parts[1]
        if not folder.is_dir():
            log.error("Folder found in DB but missing on disk: %s", folder)
            return None

        log.info("Patient folder resolved: %s", folder)
        return folder
    except Exception as e:
        log.error("DB folder lookup failed: %s", e)
        return None

def insert_document(patient, relative_path, description):
    # type: (dict, str, str) -> bool
    """Inserts a new record into PUBLIC.MDB Documents for the given patient."""
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
        log.info(
            "Insert OK: patient=%s path='%s' db=%s",
            patient["code"], relative_path, PUBLIC_MDB.name
        )
        return True
    except Exception as e:
        log.error("DB insert failed: %s", e)
        return False


# Access constant for subform control type
_AC_SUBFORM = 112


def _find_sfdoc(form):
    """Recursively finds the SFDoc subform in the control tree. Returns its Form object or None."""
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


def refresh_ui(expected_patient_code=None):
    # type: (Optional[str]) -> None
    """
    Refreshes the parent form and requeried only the SFDoc subform, then moves
    to the last record.

    Changes vs V3:
    - Change 7: accepts expected_patient_code; skips refresh if the currently
      open patient no longer matches (user navigated away during a burst).
    - Change 4: calls form.Refresh() on the parent BEFORE sfdoc.Requery() so
      that image controls repaint without moving the record pointer.
    - Change 5: clears form.Dirty before Requery() so Access does not silently
      ignore it while a record is in edit mode.
    - Change 6: wraps sfdoc.Requery() in a 3-attempt retry loop (0.5 s delay);
      falls back to sfdoc.Refresh() if all attempts fail.
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
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    "Refresh skipped: expected patient %s but current patient is %s.",
                    expected_patient_code, current_code
                )
                return

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                "Subform '%s' not found in the active form. Refresh skipped.",
                SFDOC_SUBFORM_NAME
            )
            return

        # Change 4: refresh parent form to repaint image controls (no record pointer move)
        try:
            form.Refresh()
            log.info("Refresh() on parent form (image repaint).")
        except Exception as e_pref:
            log.warning("Parent form Refresh() failed (non-blocking): %s", e_pref)

        # Change 5: dirty state guard — exit edit mode before Requery()
        try:
            if form.Dirty:
                log.info("Parent form is dirty — clearing edit mode before Requery().")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug("Could not read/clear form.Dirty: %s", e_dirty)

        # Change 6: Requery retry loop (3 attempts, 0.5 s apart; fallback to Refresh)
        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5
        requery_ok = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(
                    "Requery() on '%s' (attempt %d/%d).",
                    SFDOC_SUBFORM_NAME, attempt, _REQUERY_ATTEMPTS
                )
                requery_ok = True
                break
            except Exception as e_req:
                log.warning(
                    "Requery() attempt %d/%d failed on '%s': %s",
                    attempt, _REQUERY_ATTEMPTS, SFDOC_SUBFORM_NAME, e_req
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requery_ok:
            log.warning(
                "All Requery() attempts failed on '%s', falling back to Refresh().",
                SFDOC_SUBFORM_NAME
            )
            try:
                sfdoc.Refresh()
                log.info("Fallback Refresh() on '%s'.", SFDOC_SUBFORM_NAME)
            except Exception as e_ref:
                log.warning(
                    "Fallback Refresh() also unavailable on '%s': %s",
                    SFDOC_SUBFORM_NAME, e_ref
                )

        # Navigate to the last record so the new document is visible
        try:
            sfdoc.Recordset.MoveLast()
            log.info("MoveLast() on '%s'", SFDOC_SUBFORM_NAME)
        except Exception as e_ml:
            log.debug("MoveLast() failed on '%s': %s", SFDOC_SUBFORM_NAME, e_ml)

    except Exception as e:
        log.warning("COM refresh failed (non-blocking): %s", e)


def wait_for_file(file):
    # type: (Path) -> bool
    """Tries to open the file for reading to check if it's still locked."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug("File locked (%d/%d), retrying...", attempt, FILE_LOCK_MAX_ATTEMPTS)
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error("File still locked after %d attempts: %s", FILE_LOCK_MAX_ATTEMPTS, file)
    return False


def move_file(source, dest_folder, label=""):
    # type: (Path, Path, str) -> Optional[Path]
    """Moves the file to dest_folder, resolving name conflicts with a timestamp."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name

    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / "{0}_{1}{2}".format(source.stem, ts, source.suffix)
        log.info("Name conflict, renamed to %s", dest.name)

    try:
        shutil.move(str(source), str(dest))
        tag = "[{0}]  ".format(label) if label else ""
        log.info("%s%s -> %s", tag, source.name, dest)
        return dest
    except Exception as e:
        log.error("Move failed: %s", e)
        return None


def orphan_file(file):
    # type: (Path) -> None
    """Moves the file to the orphan folder with a warning log."""
    log.warning("Orphaning: %s", file.name)
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def worker(file_queue):
    # type: (queue.Queue) -> None
    """
    Processes files from the queue. Runs the full pipeline (lock-wait →
    patient lookup → move → DB insert) for each file, then fires a single
    UI refresh once the queue has been idle for 1.5 s (burst debounce).

    Change 7: tracks last_patient_code and passes it to refresh_ui() so the
    refresh is skipped if the operator has navigated to a different patient.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh     = False
    last_patient_code = None   # change 7: track the patient from the last insert

    try:
        while True:
            try:
                file = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)  # change 7
                    needs_refresh     = False
                    last_patient_code = None
                continue
            except Exception as e:
                log.error("Queue error: %s", e)
                continue

            log.info("Processing: %s (%d pending)", file.name, file_queue.qsize())

            if not file.exists():
                log.warning("File gone before processing: %s", file)
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error("Aborting, persistent lock: %s", file.name)
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
                        "No patient open, waiting (timeout in %d min)",
                        PATIENT_WAIT_TIMEOUT // 60
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                "Patient: %s %s (code %s)",
                patient["nom"], patient["prenom"], patient["code"]
            )

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(
                    "Could not resolve folder for patient %s. Orphaning.",
                    patient["code"]
                )
                orphan_file(file)
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            group_name    = patient_folder.parent.name
            relative_path = "\\{0}\\{1}\\{2}".format(
                group_name, patient_folder.name, dest.name
            )
            description = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]   # change 7
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)   # change 7
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    """Watchdog event handler that enqueues new image files for the worker."""

    def __init__(self, file_queue):
        # type: (queue.Queue) -> None
        super(ImageProducer, self).__init__()
        self._queue = file_queue

    def on_created(self, event):
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info("Enqueued: %s (queue size: %d)", file.name, self._queue.qsize() + 1)
        self._queue.put(file)


class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler — producer side of the producer/consumer pipeline.

    Receives creation events from the PollingObserver and enqueues image files
    on the shared queue for the worker thread.  Directory events are ignored.
    """

    def __init__(self, file_queue):
        # type: (queue.Queue) -> None
        super(ImageProducer, self).__init__()
        self._queue = file_queue

    def on_created(self, event):
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info("Enqueued: %s (queue size: %d)", file.name, self._queue.qsize() + 1)
        self._queue.put(file)


def wait_for_network_share():
    # type: () -> None
    """
    Block until SOURCE_DIR is accessible.

    Returns immediately for local drive paths (identified by a colon as the
    second character, e.g. C:\\...).  For all other paths (UNC shares, mapped
    drives without a drive letter) the function polls is_dir() every 10 seconds
    and logs a warning on each failed attempt.  The first successful check after
    at least one failure also logs an info message to confirm recovery.
    """
    source_str = str(SOURCE_DIR)
    is_unc    = source_str.startswith("\\\\") or source_str.startswith("//")
    is_local  = (
        not is_unc
        and (len(source_str) >= 2 and source_str[1] == ":")   # e.g. C:\...
    )

    if is_local:
        return

    # Network path: poll until reachable
    first_attempt = True
    while True:
        try:
            if SOURCE_DIR.is_dir():
                if not first_attempt:
                    log.info("Network share is now reachable: %s", SOURCE_DIR)
                return
        except Exception:
            pass

        log.warning(
            "Network share not reachable, retrying in 10 s: %s", SOURCE_DIR
        )
        first_attempt = False
        time.sleep(10)


def main():
    # type: () -> None

    # Change 2: wait for the network share before doing anything else
    log.info("Checking network share availability...")
    wait_for_network_share()

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 3 started")
    log.info("  Source     : %s", SOURCE_DIR)
    log.info("  Dest       : %s", DEST_PHOTOS)
    log.info("  PUBLIC.MDB : %s", PUBLIC_MDB)
    log.info("  DOCUM.MDB  : %s", DOCUM_MDB)
    log.info("  Orphans    : %s", ORPHAN_DIR)
    log.info("  Timeout    : %d min", PATIENT_WAIT_TIMEOUT // 60)
    log.info("  Ext        : %s", ", ".join(sorted(WATCHED_EXTENSIONS)))

    file_queue = queue.Queue()

    worker_thread = threading.Thread(
        target=worker, args=(file_queue,), name="Worker"
    )
    worker_thread.daemon = True
    worker_thread.start()

    # Change 3: auto-reconnect loop — restarts the Observer if the network drops
    _RECONNECT_DELAY = 15   # seconds to wait before restarting after a drop

    def _start_observer():
        # type: () -> Observer
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
                    "Observer died (possible network drop). "
                    "Waiting %d s before reconnecting...",
                    _RECONNECT_DELAY
                )
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass

                wait_for_network_share()          # block until share is back
                time.sleep(_RECONNECT_DELAY)      # extra settling pause

                log.info("Restarting Observer after network recovery.")
                observer = _start_observer()

    except KeyboardInterrupt:
        log.info("Shutdown requested.")
    finally:
        observer.stop()
        observer.join()

        remaining = file_queue.qsize()
        if remaining:
            log.info("Waiting for %d remaining file(s)...", remaining)
            file_queue.join()

        log.info("Image Router stopped.")


if __name__ == "__main__":
    main()