"""
box2.py — Image Router, Version 2 (initial Nidek-aware release)
================================================================
This module monitors a source directory for incoming medical image files
and automatically routes each one to the correct patient folder, then
registers the file in the StudioVision Access database (PUBLIC.MDB).

Processing pipeline for every detected file
-------------------------------------------
1. A Watchdog observer detects a new file in SOURCE_DIR and places it on
   a thread-safe queue (producer role).
2. A single background worker thread consumes the queue one item at a time
   (consumer role), preventing concurrent writes to the same patient folder
   and the same database table.
3. The worker waits for the file to be fully written (lock-check loop), then
   queries the currently open Access patient form via COM automation to
   identify the destination patient.
4. If no patient form is visible the worker polls every PATIENT_POLL_INTERVAL
   seconds for up to PATIENT_WAIT_TIMEOUT seconds; if the timeout expires the
   file is moved to ORPHAN_DIR.
5. The patient's folder on disk is resolved from an existing record in the
   PUBLIC.MDB Documents table.
6. The file is moved to the patient's photo folder and a new Documents row is
   inserted into PUBLIC.MDB with the appropriate description (Image / OCT /
   DICOM).
7. The Access UI is requeried immediately after the insert so the new document
   appears without a manual refresh.

Nidek device handling
---------------------
Nidek retinal cameras export a scan as a subdirectory tree:
  SOURCE_DIR/<device>/<scan_id>/<image(s)> [+ XML metadata]

For each scan directory the worker:
  - Deletes all XML metadata files (not needed by StudioVision).
  - Selects only the largest image file (the full-resolution export) and
    discards smaller thumbnails.
  - Tracks processed scan directories so that residual files still in the
    queue are silently removed rather than triggering duplicate processing.

Dependencies
------------
  - watchdog  : cross-platform filesystem event monitoring
  - pyodbc    : ODBC connection to Microsoft Access (.mdb) databases
  - pywin32   : COM automation of the running Access.Application instance
  - pythoncom : COM initialisation / teardown required per thread

Configuration
-------------
All paths (SOURCE_DIR, DEST_PHOTOS, PUBLIC_MDB, DOCUM_MDB, ORPHAN_DIR)
must be set to valid Windows UNC or local paths before deployment.
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
    """Establishes a connection to the specified Access MDB file using pyodbc.

    Args:
        mdb_path (Path): The file path of the MDB file to connect to.

    Returns:
        pyodbc.Connection: A connection object for the MDB file.
    """
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )

def get_active_patient() -> dict | None:
    """Retrieves patient information from the active Access form.

    The function checks if the expected fields are present and returns their
    values as a dictionary. If the form is not active or the fields are
    missing, None is returned.

    Returns:
        dict | None: A dictionary with patient information or None if not available.
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
    """Finds the folder where patient photos are stored based on the patient code.

    Args:
        patient_code (str): The code of the patient whose folder is to be found.

    Returns:
        Path | None: The path to the patient's folder or None if not found.
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
    """Inserts a new document record into the PUBLIC.MDB database.

    Args:
        patient (dict): A dictionary with patient information (code, nom, prenom).
        relative_path (str): The relative path of the document to be inserted.
        description (str): A description of the document (e.g., Image, OCT, DICOM).

    Returns:
        bool: True if the insert was successful, False otherwise.
    """
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert skipped.")
        return False

    # IMPORTANT: target_mdb must be PUBLIC.MDB because DOCUM.MDB is read-only for this operation
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


# Access constant for subform control type
_AC_SUBFORM = 112

def _requery_form(form) -> None:
    """Requeries the specified Access form to refresh its data.

    This function attempts to call the Requery() method on the form and falls
    back to Refresh() if Requery() is not available. It also recursively
    requeried subforms.

    Args:
        form: The Access form object to be requeried.
    """
    
    # Recurse into subforms first so their data is fresh before the parent is requeried
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
    """Moves to the last record in the document subform to display the latest document.

    This function recursively navigates through subforms if necessary.

    Args:
        form: The Access form object containing the subform.
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
    """Refreshes the active Access form to display the latest document.

    The function uses COM automation to access the running instance of
    Access.Application and attempts to requery the active form. If the form
    cannot be requeried, the function fails silently to avoid blocking the
    worker thread.

    Returns:
        None
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
    """Waits for a file to be unlocked by the writing process.

    The function retries opening the file for reading several times with a
    delay, to check if the file is still locked. If the file cannot be opened
    after the maximum number of attempts, the function logs an error.

    Args:
        file (Path): The file path to be checked.

    Returns:
        bool: True if the file is unlocked and can be opened, False otherwise.
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
    """Moves a file to the specified destination folder, renaming it if necessary.

    If a file with the same name already exists in the destination folder, the
    file is renamed by appending a timestamp to its name. The function logs the
    move operation.

    Args:
        source (Path): The source file path to be moved.
        dest_folder (Path): The destination folder where the file should be moved.
        label (str, optional): An optional label to prefix the log message. Defaults to "".

    Returns:
        Path | None: The path to the moved file in the destination folder, or None if the move failed.
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
    """Moves a file to the orphan folder and logs a warning.

    The orphan folder is used to store files that could not be processed or
    assigned to a patient. The function logs the orphaning action.

    Args:
        file (Path): The file path to be moved to the orphan folder.

    Returns:
        None
    """
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")

def _try_rmdir(folder: Path) -> None:
    """Attempts to remove a folder if it is empty.

    The function checks if the folder is a directory and if it is empty before
    attempting to remove it. If the folder cannot be removed, the function fails
    silently.

    Args:
        folder (Path): The folder path to be removed.

    Returns:
        None
    """
    try:
        if folder.is_dir() and not any(folder.iterdir()):
            folder.rmdir()
            log.info(f"Empty folder removed: {folder}")
        else:
            log.debug(f"Folder not removed (non-empty or missing): {folder}")
    except Exception as e:
        log.debug(f"_try_rmdir({folder}) ignored: {e}")

def worker(file_queue: queue.Queue) -> None:
    """The worker function that processes image files from the queue.

    This function runs in a background thread and performs the following steps for
    each file in the queue:
    - Waits for the file to be fully written and unlocked.
    - Identifies the patient by checking the active Access form.
    - Resolves the patient's photo folder from the database.
    - Moves the file to the patient's folder.
    - Inserts a new document record into the database.
    - Refreshes the Access UI to show the new document.

    Args:
        file_queue (queue.Queue): The thread-safe queue containing the files to be processed.

    Returns:
        None
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    # Tracks scan folders already processed to suppress residual files still in the queue
    processed_scan_dirs: set[Path] = set()

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

            scan_dir = file.parent
            main_dir = file.parent.parent
            is_nidek = main_dir.parent == SOURCE_DIR  

            if is_nidek:

                if scan_dir in processed_scan_dirs:
                    try:
                        file.unlink()
                        log.info(f"[NIDEK] Residual removed (scan already processed): {file.name}")
                    except Exception as e:
                        log.warning(f"[NIDEK] Could not remove residual {file.name}: {e}")

                    _try_rmdir(scan_dir)
                    _try_rmdir(main_dir)
                    if not scan_dir.exists():
                        processed_scan_dirs.discard(scan_dir)

                    file_queue.task_done()
                    continue

                log.info(f"[NIDEK] Stabilising '{scan_dir.name}' (parent: '{main_dir.name}')...")
                time.sleep(2)

                for xml_file in list(scan_dir.glob("*.xml")):
                    try:
                        xml_file.unlink()
                        log.info(f"[NIDEK] XML removed: {xml_file.name}")
                    except Exception as e:
                        log.warning(f"[NIDEK] Could not remove {xml_file.name}: {e}")


                sibling_images = [
                    f for f in scan_dir.iterdir()
                    if f.is_file() and f.suffix.lower() in WATCHED_EXTENSIONS
                ]

                if not sibling_images:
                    log.warning(f"[NIDEK] No images found in '{scan_dir.name}', skipping.")
                    file_queue.task_done()
                    continue

                largest_image = max(sibling_images, key=lambda f: f.stat().st_size)

                if file.resolve() != largest_image.resolve():

                    try:
                        file.unlink()
                        log.info(f"[NIDEK] Thumbnail removed: {file.name}")
                    except Exception as e:
                        log.warning(f"[NIDEK] Could not remove thumbnail {file.name}: {e}")
                    file_queue.task_done()
                    continue

                log.info(f"[NIDEK] Main image identified: {file.name} "
                         f"({file.stat().st_size:,} bytes)")


                processed_scan_dirs.add(scan_dir)

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
                    log.info(f"No patient open, waiting "
                             f"(timeout in {PATIENT_WAIT_TIMEOUT // 60} min)")
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
            
            if is_nidek:
                _try_rmdir(scan_dir)
                _try_rmdir(main_dir)
                if not scan_dir.exists():
                    processed_scan_dirs.discard(scan_dir)
                    log.info(f"[NIDEK] scan_dir cleared from tracking set.")

            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
                time.sleep(1.5)
                refresh_ui()
            else:
                log.warning("Insert failed, refresh skipped.")

            file_queue.task_done()

    finally:
        pythoncom.CoUninitialize()

class ImageProducer(FileSystemEventHandler):
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
    """Main function to start the image routing process.

    This function initializes the source and destination directories, sets up
    the orphan directory, and starts the Watchdog observer and the worker
    thread. It also logs the initial configuration and waits for the observer
    to stop on exit.

    Returns:
        None
    """
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