"""
studiovision_monitorV2.py
Image Router — Version 2.

Introduces the following improvements over the Version 1 baseline:

  1. Burst debounce — instead of calling Access COM immediately after each
     successful DB insert, the worker raises a needs_refresh flag and defers
     the refresh until the queue has been idle for 1.5 seconds.  This reduces
     COM round-trips during rapid multi-file acquisitions and avoids freezing
     the Access UI between consecutive inserts.

  2. Targeted SFDoc-only refresh — instead of requerying the entire parent
     form (which resets its record pointer to #1), the function now locates
     the SFDoc subform directly via _find_sfdoc() and calls Requery() only on
     it, followed by MoveLast() to position the cursor on the new document.

  3. Log file in ~/studiovision/ — ensures the log path is valid both when the
     script is run directly and when distributed as a compiled executable.

Architecture is the same producer/consumer single-worker-thread design as the
baseline, using the standard Watchdog Observer (not polling-based, so suited
for local or already-connected network drives only).

Dependencies: watchdog, pyodbc, pywin32, pythoncom
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

# Configure logging to file and console with timestamps and thread names
import os
_log_path = os.path.join(os.path.expanduser("~"), "studiovision", "image_router.log")
os.makedirs(os.path.dirname(_log_path), exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_log_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")


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


def refresh_ui() -> None:
    """Requeried only the SFDoc subform and moves to the last record. 
    Never touches the parent form. All COM errors are caught and logged."""
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found in the active form. "
                "Refresh skipped."
            )
            return

        # Requery the subform only, leaving the parent form's record pointer unchanged.
        try:
            sfdoc.Requery()
            log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_req:
            log.warning(
                f"Requery() unavailable on '{SFDOC_SUBFORM_NAME}' ({e_req}), "
                "trying Refresh()..."
            )
            try:
                sfdoc.Refresh()
                log.info(f"Refresh() on '{SFDOC_SUBFORM_NAME}'")
            except Exception as e_ref:
                log.warning(
                    f"Refresh() also unavailable on '{SFDOC_SUBFORM_NAME}' ({e_ref})"
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
    Long-running consumer thread that drains the shared queue and processes
    each image through the complete pipeline.

    COM is initialised per-thread via CoInitialize()/CoUninitialize().

    Burst debounce (improvement 1): the worker uses a 1.5-second blocking
    get() timeout.  While files arrive faster than that interval it processes
    them back-to-back with needs_refresh=True.  Once the queue is idle for
    1.5 seconds (burst end), a single UI refresh is fired rather than one
    after every individual insert, avoiding repeated Access COM round-trips.

    Processing pipeline per file:
      1. Existence check — the file may be deleted before the worker dequeues it.
      2. Lock-wait — polls until the file is fully written (see wait_for_file).
      3. Patient identification loop — polls the active Access form until a
         patient is found or PATIENT_WAIT_TIMEOUT elapses; orphans on timeout.
      4. Folder resolution — queries PUBLIC.MDB for the patient's directory.
      5. File move — transfers the image into the patient folder.
      6. DB insert — registers the image in the Documents table.
      7. Burst flag — sets needs_refresh=True for the deferred UI update.

    On shutdown any pending refresh is flushed in the finally block.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh: bool = False

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                # Queue has been idle for 1.5 s — the burst is over.  Fire the
                # deferred refresh exactly once rather than after every insert.
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui()
                    needs_refresh = False
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
                # Raise the burst flag — the refresh fires once the queue is
                # idle, not immediately, to avoid freezing Access on rapid
                # sequential inserts.
                needs_refresh = True
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")

            file_queue.task_done()

    finally:
        # Flush any pending refresh before the thread exits
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui()
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    """Watchdog event handler that enqueues new image files for the worker."""

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
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 2 started")
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