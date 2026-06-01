"""
Routes incoming imaging files to the correct patient folder,
inserts a DB record, and refreshes the Access UI.

Pipeline: PollingObserver → file_queue → Worker → Access DB + UI refresh
          (1.5s burst debounce, auto-reconnect on network drop)

New in v5:
  - Startup scan + independent catchup loop
  - Progressive retry queue on DB failure
  - Smart post-commit delay before Requery/MoveLast
  - Simplified doctor log in French (rapport_medecin_YYYY-MM-DD.txt)

Dependencies: watchdog, pyodbc, pywin32, pythoncom, pystray, Pillow, psutil
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
import subprocess
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

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = False
except ImportError:
    TRAY_AVAILABLE = False

import win32api
import win32event
import winerror
import psutil

# Configuration
BOX_NAME    = "Studiovision"

STUDIO_VISION_EXE = "studiovision.exe"

SOURCE_DIR  = Path(r"C:\Users\Box-6\Desktop\OM")
ORPHAN_DIR  = Path(r"C:\Users\Box-6\Desktop\Images_Oubliées")
DEST_PHOTOS = Path(r"\\studiovision\Studiov2000-OM\PHOTOS")
PUBLIC_MDB  = Path(r"\\studiovision\Studiov2000-OM\PUBLIC.MDB")
DOCUM_MDB   = Path(r"\\studiovision\Studiov2000-OM\DOCUM.MDB")

WATCHED_EXTENSIONS     = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff", ".dcm", ".pdf", ".rtf", ".doc", ".docx", ".odt"}
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3
PATIENT_WAIT_TIMEOUT   = 900

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

SFDOC_SUBFORM_NAME = "SFDoc"

EXAM_DESCRIPTION = {
    ".jpg":  "Champ visuel",
    ".jpeg": "Champ visuel",
    ".jfif": "Champ visuel",
    ".png":  "Champ visuel",
    ".bmp":  "Champ visuel",
    ".tif":  "Champ visuel",
    ".tiff": "Champ visuel",
    ".dcm":  "Champ visuel",
    ".pdf":  "Champ visuel",
    ".rtf":  "Champ visuel",
    ".doc":  "Champ visuel",
    ".docx": "Champ visuel",
    ".odt":  "Champ visuel",
}

# Progressive retry delays (seconds) between DB insert attempts
RETRY_DELAYS = [10, 30, 60, 120, 300]
# Interval (seconds) between catchup scans of the source folder
CATCHUP_INTERVAL = 120

# Logging
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

# Doctor log (simplified, in French)

def _get_medecin_log_path() -> Path:
    """Returns the path to today's doctor report."""
    date_str = datetime.now().strftime("%Y-%m-%d")
    return _LOG_DIR / f"rapport_medecin_{date_str}.txt"


def log_medecin(message: str) -> None:
    """Appends a timestamped line to the daily doctor report."""
    heure = datetime.now().strftime("%Hh%M")
    ligne = f"[{heure}] {message}\n"
    try:
        with open(_get_medecin_log_path(), "a", encoding="utf-8") as f:
            f.write(ligne)
    except Exception as e:
        log.warning(f"Could not write to doctor report: {e}")


# System tray
_NETWORK_SHARE_POLL = 10

_ICON_SIZE     = 64
_COLOR_READY   = (30, 144, 255)
_COLOR_ACTIVE  = (50, 205, 50)

_icon: "pystray.Icon | None" = None
_status_text: str             = "Starting..."
_stop_event: threading.Event  = threading.Event()

_mutex_handle = None


def _make_icon(color: tuple) -> "Image.Image":
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse(
        [margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin],
        fill=color,
    )
    return img


def _set_status(text: str, processing: bool = False) -> None:
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Tray update failed: {e}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            log.debug(f"Notification failed: {e}")


def _open_logs(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning(f"Could not open log file: {e}")


def _open_medecin_log(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_get_medecin_log_path()))
    except Exception as e:
        log.warning(f"Could not open doctor report: {e}")


def _quit(icon, item) -> None:  # noqa: ARG001
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


# Network
def wait_for_network_share() -> None:
    is_network = str(SOURCE_DIR).startswith("\\\\") or str(SOURCE_DIR).startswith("//")
    if not is_network:
        return
    attempt = 0
    while not SOURCE_DIR.is_dir():
        attempt += 1
        log.warning(
            f"Network share not reachable: {SOURCE_DIR}  "
            f"(attempt {attempt}, retrying in {_NETWORK_SHARE_POLL}s)"
        )
        time.sleep(_NETWORK_SHARE_POLL)
    if attempt:
        log.info(f"Network share accessible after {attempt} attempt(s): {SOURCE_DIR}")


# Access database
def db_connect(mdb_path: Path):
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


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


def insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """Inserts a document record into the DB. Returns True on success."""
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


# Access UI refresh
_AC_SUBFORM = 112

# Delay (seconds) between commit() and Requery/MoveLast to allow Jet/ACE to finalize the write.
_UI_POST_COMMIT_DELAY = 0.5


def _find_sfdoc(form):
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


def refresh_ui(expected_patient_code: str | None = None) -> None:
    """
    Refreshes the Access UI after an INSERT.
    Waits _UI_POST_COMMIT_DELAY before Requery to let Jet/ACE finalize the write,
    then calls MoveLast() to select the newly inserted record.
    """
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        if expected_patient_code is not None:
            current = get_active_patient()
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    f"Refresh skipped: active patient changed "
                    f"(expected={expected_patient_code}, current={current_code})."
                )
                return

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

        try:
            if form.Dirty:
                log.info("Parent form is dirty; clearing before Requery.")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Dirty check/clear failed ({e_dirty}), continuing...")

        log.debug(f"Waiting {_UI_POST_COMMIT_DELAY}s before Requery (post-commit settle)...")
        time.sleep(_UI_POST_COMMIT_DELAY)

        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5

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

        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}' — last record selected.")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


# File operations
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


def orphan_file(file: Path) -> None:
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def prevent_sleep() -> None:
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning(f"Could not set execution state: {e}")


# Retry queue for deferred DB inserts
# Each entry: { "file", "patient", "relative_path", "description", "attempt", "next_try" }
_retry_queue: list[dict] = []
_retry_lock  = threading.Lock()


def _schedule_retry(file: Path, patient: dict, relative_path: str,
                    description: str, attempt: int = 0) -> None:
    """Schedules a failed insert for retry with progressive delay. Orphans the file after max attempts."""
    if attempt >= len(RETRY_DELAYS):
        log.error(
            f"Max retry attempts reached for '{file.name}' "
            f"(patient {patient['code']}). Orphaning."
        )
        log_medecin(
            f"ERREUR : Impossible d'ajouter la photo '{file.name}' après "
            f"{len(RETRY_DELAYS)} tentatives (Réseau déconnecté ou base inaccessible)."
        )
        orphan_file(file)
        return

    delay      = RETRY_DELAYS[attempt]
    next_try   = time.monotonic() + delay
    entry      = {
        "file":          file,
        "patient":       patient,
        "relative_path": relative_path,
        "description":   description,
        "attempt":       attempt,
        "next_try":      next_try,
    }
    with _retry_lock:
        _retry_queue.append(entry)

    log.warning(
        f"DB insert failed for '{file.name}' — "
        f"scheduled retry #{attempt + 1} in {delay}s."
    )
    log_medecin(
        f"AVERTISSEMENT : Échec temporaire pour la photo '{file.name}' "
        f"(patient {patient['nom']} {patient['prenom']}, Code: {patient['code']}). "
        f"Nouvelle tentative dans {delay} secondes."
    )


def _process_retry_queue(file_queue: queue.Queue) -> None:
    """Processes due retry entries. On success triggers a UI refresh; on failure reschedules."""
    now = time.monotonic()
    to_retry: list[dict] = []

    with _retry_lock:
        remaining = []
        for entry in _retry_queue:
            if entry["next_try"] <= now:
                to_retry.append(entry)
            else:
                remaining.append(entry)
        _retry_queue[:] = remaining

    for entry in to_retry:
        file          = entry["file"]
        patient       = entry["patient"]
        relative_path = entry["relative_path"]
        description   = entry["description"]
        attempt       = entry["attempt"]

        log.info(
            f"Retry #{attempt + 1} for '{file.name}' "
            f"(patient {patient['code']})..."
        )

        if insert_document(patient, relative_path, description):
            log.info(f"Retry #{attempt + 1} succeeded for '{file.name}'.")
            log_medecin(
                f"SUCCÈS (rattrapage) : Photo '{file.name}' ajoutée au dossier de "
                f"{patient['nom']} {patient['prenom']} (Code: {patient['code']})."
            )
            _notify("Rattrapage OK", f"{file.name} inséré après {attempt + 1} tentative(s)")
            refresh_ui(expected_patient_code=patient["code"])
        else:
            _schedule_retry(file, patient, relative_path, description, attempt + 1)


# Startup scan and periodic catchup
def _scan_source_for_missed_files(file_queue: queue.Queue) -> None:
    """Scans SOURCE_DIR for files missed during downtime and enqueues them."""
    if not SOURCE_DIR.is_dir():
        log.warning(f"Catchup scan skipped: {SOURCE_DIR} not accessible.")
        return

    found = []
    try:
        for item in SOURCE_DIR.rglob("*"):
            if item.is_file() and item.suffix.lower() in WATCHED_EXTENSIONS:
                found.append(item)
    except Exception as e:
        log.error(f"Error during catchup scan: {e}")
        return

    if not found:
        log.info("Catchup scan: no pending files.")
        return

    log.info(f"Catchup scan: {len(found)} file(s) found, enqueuing.")
    for f in found:
        file_queue.put(f)
        log.info(f"  Catchup → enqueued: {f.name}")


def _catchup_loop(file_queue: queue.Queue) -> None:
    """Periodically scans SOURCE_DIR and processes the retry queue."""
    log.info(f"Catchup thread started (interval: {CATCHUP_INTERVAL}s).")
    while not _stop_event.is_set():
        for _ in range(CATCHUP_INTERVAL):
            if _stop_event.is_set():
                break
            time.sleep(1)

        if _stop_event.is_set():
            break

        log.debug("Catchup: scanning SOURCE_DIR + processing retry queue...")
        _scan_source_for_missed_files(file_queue)
        _process_retry_queue(file_queue)

    log.info("Catchup thread stopped.")


# Worker (queue consumer)
def worker(file_queue: queue.Queue) -> None:
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh: bool           = False
    last_patient_code: str | None = None
    burst_count: int              = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh = False
                    last_patient_code = None
                    _notify(
                        "Transfer complete",
                        f"{burst_count} file(s) processed",
                    )
                    _set_status(f"{BOX_NAME} — Ready", processing=False)
                    burst_count = 0
                continue
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue

            log.info(f"Processing: {file.name} ({file_queue.qsize()} pending)")

            if burst_count == 0 and not needs_refresh:
                _notify("Transfer in progress", file.name)
            _set_status("Transfer in progress...", processing=True)

            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                _notify("Error", f"File locked: {file.name}")
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
                    _notify("Orphan file", file.name)
                    log_medecin(
                        f"AVERTISSEMENT : Fichier '{file.name}' orphelin "
                        f"(aucun patient ouvert après {PATIENT_WAIT_TIMEOUT // 60} min)."
                    )
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
                _notify("Orphan file", file.name)
                log_medecin(
                    f"ERREUR : Dossier introuvable pour '{file.name}' "
                    f"(patient {patient['nom']} {patient['prenom']}, "
                    f"Code: {patient['code']})."
                )
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
                burst_count      += 1
                log.debug("Insert OK — refresh deferred to burst end.")
                log_medecin(
                    f"SUCCÈS : Photo '{dest.name}' ajoutée au dossier de "
                    f"{patient['nom']} {patient['prenom']} (Code: {patient['code']})."
                )
            else:
                log.warning("Insert failed — scheduling retry.")
                _notify("DB Error", f"Insert failed — retry scheduled for {file.name}")
                _schedule_retry(dest, patient, relative_path, description, attempt=0)

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                _notify("Transfer complete", f"{burst_count} file(s) processed")
        _set_status(f"{BOX_NAME} — Stopped")
        pythoncom.CoUninitialize()


# Watchdog producer
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


# Background thread (watchdog + network reconnect)
def _run_background(file_queue: queue.Queue) -> None:
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — watching for images.")
        return obs

    observer = _start_observer()
    _set_status(f"{BOX_NAME} — Ready", processing=False)

    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observer stopped (network drop?). Attempting reconnect...")
                _set_status(f"{BOX_NAME} — Reconnecting...", processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                log.info(f"Waiting {_RECONNECT_WAIT}s before restarting observer...")
                time.sleep(_RECONNECT_WAIT)
                observer = _start_observer()
                _set_status(f"{BOX_NAME} — Ready", processing=False)
            time.sleep(1)
    finally:
        observer.stop()
        observer.join()
        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Waiting for {remaining} remaining file(s)...")
            file_queue.join()
        log.info("Background thread stopped.")
        if _icon is not None:
            _icon.stop()


# Entry point
def main() -> None:
    global _icon, _mutex_handle

    # Single-instance guard
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

    # Block manual double-click restart while Studio Vision is running
    try:
        parent_name = psutil.Process(os.getpid()).parent().name().lower()
    except Exception:
        parent_name = ""

    if parent_name == "explorer.exe":
        sv_running = any(
            (p.info["name"] or "").lower() == STUDIO_VISION_EXE
            for p in psutil.process_iter(["name"])
        )
        if sv_running:
            ctypes.windll.user32.MessageBoxW(
                0,
                "Pour relancer le routeur d'images, veuillez fermer complètement "
                "puis relancer Studio Vision.",
                "Routeur d'images",
                0x30,  # MB_ICONWARNING | MB_OK
            )
            sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 5 started")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Dest       : {DEST_PHOTOS}")
    log.info(f"  PUBLIC.MDB : {PUBLIC_MDB}")
    log.info(f"  DOCUM.MDB  : {DOCUM_MDB}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    log.info(f"  Log file   : {_LOG_FILE}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext        : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info(f"  Catchup    : every {CATCHUP_INTERVAL}s")
    log.info(f"  Retry      : {RETRY_DELAYS}")
    log.info(f"  Doctor log : {_get_medecin_log_path()}")

    log_medecin(
        f"Démarrage du routeur d'images (version 5). "
        f"Surveillance de : {SOURCE_DIR}"
    )

    file_queue: queue.Queue = queue.Queue()

    # 1. Worker (consumer)
    threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    ).start()

    # 2. Startup scan: catch files dropped during downtime
    log.info("Startup scan — checking for pending files...")
    _scan_source_for_missed_files(file_queue)

    # 3. Watchdog thread (main producer)
    threading.Thread(
        target=_run_background, args=(file_queue,), name="Background", daemon=True
    ).start()

    # 4. Catchup thread (periodic scan + retry)
    threading.Thread(
        target=_catchup_loop, args=(file_queue,), name="Catchup", daemon=True
    ).start()

    if not TRAY_AVAILABLE:
        log.info("Mode invisible activé. Lancement de Studio Vision OM...")
        
        # La commande exacte pour lancer ta base OM (récupérée de ton premier script)
        cmd_om = [
            r"C:\Studiov2000-OM\svprog\msaccess.exe",
            "/runtime", r"C:\Studiov2000-OM\svprog\Ophprog.mde",
            "/wrkgrp",  r"C:\Studiov2000-OM\config\system.mdw",
            "/User", "/Pwd", "/X", "demarrage"
        ]

        try:
            # On lance Studio Vision. 
            # subprocess.run met le script Python en pause sur cette ligne tant que le logiciel est ouvert !
            subprocess.run(cmd_om, check=False)
            log.info("Studio Vision OM a été fermé par l'utilisateur.")
        except FileNotFoundError:
            log.error("Erreur : L'exécutable msaccess.exe ou Studio Vision est introuvable.")
        except Exception as e:
            log.error(f"Erreur lors du lancement de Studio Vision OM : {e}")
        finally:
            # Quand le logiciel se ferme (ou plante), on arrête proprement notre script Python
            _stop_event.set()
            log.info("Arrêt du moniteur OM synchronisé avec la fermeture du logiciel.")
            
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Ouvrir le log technique", _open_logs),
        pystray.MenuItem("Ouvrir le rapport médecin", _open_medecin_log),
        pystray.MenuItem("Quitter", _quit),
    )

    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=_make_icon(_COLOR_READY),
        title=BOX_NAME,
        menu=menu,
    )

    log.info("System tray icon started.")
    _icon.run()

    _stop_event.set()
    log.info("Application stopped.")
    log_medecin("Arrêt du routeur d'images.")

if __name__ == "__main__":
    main()