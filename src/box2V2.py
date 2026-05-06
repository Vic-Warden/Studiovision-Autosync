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

# Helper to connect to an Access MDB with pyodbc, with error handling deferred to caller
def db_connect(mdb_path: Path):
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )

# Returns a dict with patient info if an Access form with the expected fields is active, else None
def get_active_patient() -> dict | None:
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

# Uses the PUBLIC.MDB Documents table to resolve the patient's photo folder,
# returning a Path if successful or None if any step fails
def find_patient_folder(patient_code: str) -> Path | None:
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

# Inserts a new record into PUBLIC.MDB Documents for the given patient, relative path, and description.
def insert_document(patient: dict, relative_path: str, description: str) -> bool:
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


def _find_sfdoc(form):
    """
    Recursively walks the form's control tree and returns the Form object
    of the subform named SFDOC_SUBFORM_NAME, or None if not found.
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
            pass
    return None


def refresh_ui() -> None:
    """
    Two-step refresh strategy:
      1. Refresh() on the PARENT form — updates bound image controls (photo
         display) without resetting its current-record pointer. We use
         Refresh() instead of Requery() on the parent to avoid jumping back
         to record #1.
      2. Requery() + MoveLast() on SFDoc only — reloads the document list
         and positions it on the newly added entry.
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

        # --- Step 1: Refresh the parent form to update image controls ---
        # Refresh() repaints bound controls from the current record without
        # moving the record pointer, so the consultation stays in place.
        try:
            form.Refresh()
            log.info(f"Refresh() on parent form '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() on parent form failed ({e_ref}), continuing...")

        # --- Step 2: Requery SFDoc to load the new document row ---
        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found in the active form. "
                "SFDoc refresh skipped."
            )
            return

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
            except Exception as e_ref2:
                log.warning(
                    f"Refresh() also unavailable on '{SFDOC_SUBFORM_NAME}' ({e_ref2})"
                )

        # Navigate to the last record so the new document is visible
        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


# Tries to open the file for reading to check if it's still locked by the writing process
def wait_for_file(file: Path) -> bool:
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retrying...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error(f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {file}")
    return False

# Moves the file to the destination folder, handling name conflicts by appending a timestamp
def move_file(source: Path, dest_folder: Path, label: str = "") -> Path | None:
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

# Moves the file to the orphan folder with a warning log
def orphan_file(file: Path) -> None:
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")

# Removes a folder only if it is strictly empty; fails silently otherwise
def _try_rmdir(folder: Path) -> None:
    try:
        if folder.is_dir() and not any(folder.iterdir()):
            folder.rmdir()
            log.info(f"Empty folder removed: {folder}")
        else:
            log.debug(f"Folder not removed (non-empty or missing): {folder}")
    except Exception as e:
        log.debug(f"_try_rmdir({folder}) ignored: {e}")

# Worker thread function that processes files from the queue, with patient lookup, file moving,
# DB insertion, and UI refresh logic
def worker(file_queue: queue.Queue) -> None:
    pythoncom.CoInitialize()
    log.info("Worker started.")

    # Tracks scan folders already processed to suppress residual files still in the queue
    processed_scan_dirs: set[Path] = set()

    needs_refresh: bool = False

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
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

            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
                needs_refresh = True
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")

            # Clean up scan folders after file transfer and DB insert
            if is_nidek:
                _try_rmdir(scan_dir)
                _try_rmdir(main_dir)
                if not scan_dir.exists():
                    processed_scan_dirs.discard(scan_dir)
                    log.info(f"[NIDEK] scan_dir cleared from tracking set.")

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui()
        pythoncom.CoUninitialize()


# Watchdog event handler that enqueues new image files for processing by the worker thread
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

# Main function to start the image router
def main() -> None:
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Image Router v3.8 started")
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