"""
Routes incoming imaging files to the correct patient folder,
inserts a DB record, and refreshes the Access UI.

Windows 7 / Python 3.8.10 compatible.

Pipeline : PollingObserver → file_queue → Worker → Access DB + UI refresh
           (1.5 s burst debounce, auto-reconnect on network drop)
Catch-up  : Independent sweep thread re-enqueues any file stranded in SOURCE_DIR.

Dependencies: watchdog, pyodbc, pywin32, pythoncom, pystray, Pillow, psutil
"""

import ctypes
import logging
import logging.handlers
import os
import pythoncom
import queue
import shutil
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, Set

import win32api
import win32event
import winerror
import psutil

from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver as Observer

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
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False


 
# Configuration
 

BOX_NAME          = "Windows 7"
STUDIO_VISION_EXE = "studiovision.exe"

SOURCE_DIR  = Path(r"??")
ORPHAN_DIR  = Path(r"??")
DEST_PHOTOS = Path(r"??")
PUBLIC_MDB  = Path(r"??")
DOCUM_MDB   = Path(r"??")

WATCHED_EXTENSIONS = {
    ".jpg", ".jpeg", ".jfif", ".png", ".bmp",
    ".tif", ".tiff", ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3
PATIENT_WAIT_TIMEOUT   = 900

# How often the catch-up sweep runs (seconds).
SWEEP_INTERVAL_SECONDS = 300  # 5 minutes

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"
SFDOC_SUBFORM_NAME  = "SFDoc"
_AC_SUBFORM         = 112   # Access ControlType constant for subform

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
    ".odt":  "Document",
}


 
# Logging — two handlers per day
#   • image_router_YYYY-MM-DD.log  : full technical log
#   • transferts_medecin.log       : plain-French summary for end-users
 

_LOG_DIR = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)

_LOG_FILE_TECH   = os.path.join(_LOG_DIR, "image_router.log")
_LOG_FILE_MEDECIN = os.path.join(_LOG_DIR, "transferts_medecin.log")

# --- Technical logger (daily rotation, keep 30 days) ---
_tech_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_TECH,
    when="midnight",
    interval=1,
    backupCount=30,
    encoding="utf-8",
)
_tech_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)

_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )
)

logging.basicConfig(level=logging.INFO, handlers=[_tech_handler, _console_handler])
log = logging.getLogger("image_router")

# --- Doctor-facing logger (plain French, daily rotation, keep 90 days) ---
_medecin_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_MEDECIN,
    when="midnight",
    interval=1,
    backupCount=90,
    encoding="utf-8",
)
_medecin_handler.setFormatter(logging.Formatter("%(message)s"))

_medecin_log = logging.getLogger("medecin")
_medecin_log.setLevel(logging.INFO)
_medecin_log.propagate = False          # keep it separate from the technical log
_medecin_log.addHandler(_medecin_handler)


def _log_medecin(msg: str) -> None:
    """Write a timestamped, plain-French line to the doctor-facing log."""
    timestamp = datetime.now().strftime("%H:%M")
    _medecin_log.info("%s - %s", timestamp, msg)


 
# System tray
 

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)   # dodger blue
_COLOR_ACTIVE = (50, 205, 50)    # lime green

_icon         = None              # type: Optional[pystray.Icon]
_status_text  = "Démarrage..."    # type: str
_stop_event   = threading.Event()
_mutex_handle = None


def _make_icon(color: tuple) -> "Image.Image":
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    m    = 4
    draw.ellipse([m, m, _ICON_SIZE - m, _ICON_SIZE - m], fill=color)
    return img


def _set_status(text: str, processing: bool = False) -> None:
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as exc:
            log.debug("Tray update failed: %s", exc)


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as exc:
            log.debug("Notification failed: %s", exc)


def _open_logs(icon: object, item: object) -> None:
    try:
        os.startfile(_LOG_DIR)
    except Exception as exc:
        log.warning("Could not open log folder: %s", exc)


def _quit(icon: object, item: object) -> None:
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


 
# Database helpers
 

def db_connect(mdb_path: Path) -> "pyodbc.Connection":
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(mdb_path) + ";"
    )


def get_active_patient() -> Optional[dict]:
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
                name = str(ctrl.Name)
                if name in target:
                    data[name] = ctrl.Value
            except Exception:
                pass

        if not target.issubset(data.keys()):
            return None

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }

    except Exception as exc:
        log.debug("COM error in get_active_patient: %s", exc)
        return None


def find_patient_folder(patient_code: str) -> Optional[Path]:
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
            (int(patient_code),),
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

    except Exception as exc:
        log.error("DB folder lookup failed: %s", exc)
        return None


def insert_document(patient: dict, relative_path: str, description: str) -> bool:
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
            (int(patient["code"]), datetime.now(), description, relative_path, relative_path),
        )
        conn.commit()
        conn.close()
        log.info(
            "Insert OK: patient=%s path='%s' db=%s",
            patient["code"], relative_path, PUBLIC_MDB.name,
        )
        return True

    except Exception as exc:
        log.error("DB insert failed: %s", exc)
        return False


 
# Access UI refresh — deterministic MoveLast()
 

def _find_sfdoc(form: object) -> Optional[object]:
    """Recursively search for the SFDoc subform in the control tree."""
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


def refresh_ui(expected_patient_code: Optional[str] = None) -> None:
    """
    Requery the SFDoc subform and move the cursor to the last (newest) record.

    Determinism improvements over the previous version:
    - Forces Access window to the foreground so COM calls land on a visible,
      focused window (avoids silent failures on minimised/background windows).
    - Waits briefly after Requery() before calling MoveLast() so the recordset
      is fully populated before navigation.
    - Falls back to a DoCmd.RunCommand acCmdRecordsGoToLast when COM Recordset
      navigation fails (Access internal command, bypasses recordset state).
    - All COM exceptions are caught individually and logged, never swallowed.
    """
    if not WIN32_AVAILABLE:
        return

    _REQUERY_ATTEMPTS   = 3
    _REQUERY_DELAY      = 0.5   # seconds between retries
    _POST_REQUERY_PAUSE = 0.3   # let Access repopulate the recordset

    try:
        access = win32com.client.GetActiveObject("Access.Application")

        # --- Bring Access to the foreground so COM calls hit an active window ---
        try:
            hwnd = access.hWndAccessApp()
            if hwnd:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                time.sleep(0.1)
        except Exception as exc:
            log.debug("Could not bring Access to foreground: %s", exc)

        form = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        # --- Patient-code guard: abort if user switched patient ---
        if expected_patient_code is not None:
            current      = get_active_patient()
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    "Refresh skipped: expected patient %s but current is %s.",
                    expected_patient_code, current_code,
                )
                return

        # --- Clear dirty state on parent form to unblock Requery ---
        try:
            if form.Dirty:
                log.info("Parent form is dirty — clearing edit mode before Requery().")
                form.Dirty = False
        except Exception as exc:
            log.debug("Could not read/clear form.Dirty: %s", exc)

        # --- Repaint parent form (image control refresh) ---
        try:
            form.Refresh()
            log.info("Refresh() on parent form OK.")
        except Exception as exc:
            log.warning("Parent form Refresh() failed (non-blocking): %s", exc)

        # --- Locate SFDoc subform ---
        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                "Subform '%s' not found in the active form. Refresh skipped.",
                SFDOC_SUBFORM_NAME,
            )
            return

        # --- Requery with retries ---
        requery_ok = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(
                    "Requery() on '%s' OK (attempt %d/%d).",
                    SFDOC_SUBFORM_NAME, attempt, _REQUERY_ATTEMPTS,
                )
                requery_ok = True
                break
            except Exception as exc:
                log.warning(
                    "Requery() attempt %d/%d failed on '%s': %s",
                    attempt, _REQUERY_ATTEMPTS, SFDOC_SUBFORM_NAME, exc,
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requery_ok:
            log.warning(
                "All Requery() attempts failed on '%s' — falling back to Refresh().",
                SFDOC_SUBFORM_NAME,
            )
            try:
                sfdoc.Refresh()
                log.info("Fallback Refresh() on '%s' OK.", SFDOC_SUBFORM_NAME)
            except Exception as exc:
                log.warning("Fallback Refresh() also failed on '%s': %s", SFDOC_SUBFORM_NAME, exc)

        # --- Give Access time to populate the refreshed recordset ---
        time.sleep(_POST_REQUERY_PAUSE)

        # --- MoveLast(): primary path via Recordset, then DoCmd fallback ---
        move_ok = False
        try:
            rs = sfdoc.Recordset
            if rs is not None and rs.RecordCount > 0:
                rs.MoveLast()
                sfdoc.Bookmark = rs.Bookmark  # sync the form cursor to the recordset
                log.info("MoveLast() via Recordset on '%s' OK.", SFDOC_SUBFORM_NAME)
                move_ok = True
            else:
                log.debug(
                    "Recordset on '%s' is None or empty — skipping MoveLast().",
                    SFDOC_SUBFORM_NAME,
                )
                move_ok = True  # empty recordset is not an error
        except Exception as exc:
            log.warning(
                "MoveLast() via Recordset failed on '%s': %s — trying DoCmd fallback.",
                SFDOC_SUBFORM_NAME, exc,
            )

        if not move_ok:
            # DoCmd.RunCommand 505 = acCmdRecordsGoToLast
            # This fires the built-in Access command regardless of recordset state.
            try:
                sfdoc.SetFocus()
                access.DoCmd.RunCommand(505)
                log.info("MoveLast() via DoCmd(505) on '%s' OK.", SFDOC_SUBFORM_NAME)
            except Exception as exc:
                log.warning(
                    "DoCmd fallback for MoveLast() also failed on '%s': %s",
                    SFDOC_SUBFORM_NAME, exc,
                )

    except Exception as exc:
        log.warning("COM refresh failed (non-blocking): %s", exc)


 
# File-system utilities
 

def wait_for_file(file: Path) -> bool:
    """Block until the file is readable (not locked by another process)."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug("File locked (%d/%d), retrying...", attempt, FILE_LOCK_MAX_ATTEMPTS)
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error("File still locked after %d attempts: %s", FILE_LOCK_MAX_ATTEMPTS, file)
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> Optional[Path]:
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / "{0}_{1}{2}".format(source.stem, ts, source.suffix)
        log.info("Name conflict — renamed to %s", dest.name)
    try:
        shutil.move(str(source), str(dest))
        tag = "[{0}]  ".format(label) if label else ""
        log.info("%s%s -> %s", tag, source.name, dest)
        return dest
    except Exception as exc:
        log.error("Move failed: %s", exc)
        return None


def orphan_file(file: Path) -> None:
    log.warning("Orphaning: %s", file.name)
    move_file(file, ORPHAN_DIR, label="ORPHAN")
    _log_medecin(
        "Fichier non attribué (aucun patient ouvert) : {0} — déplacé dans le dossier orphelins.".format(
            file.name
        )
    )


def wait_for_network_share() -> None:
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
                    log.info("Network share is now reachable: %s", SOURCE_DIR)
                return
        except Exception:
            pass
        log.warning("Network share not reachable, retrying in 10 s: %s", SOURCE_DIR)
        first_attempt = False
        time.sleep(10)


def prevent_sleep() -> None:
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as exc:
        log.warning("Could not set execution state: %s", exc)


 
# Worker thread — consumer
 

def worker(file_queue: queue.Queue) -> None:
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh:     bool           = False
    last_patient_code: Optional[str]  = None
    burst_count:       int            = 0

    try:
        while True:
            # --- Debounce: if nothing arrives within 1.5 s, flush pending refresh ---
            try:
                file = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh     = False
                    last_patient_code = None
                    _notify(
                        "Transfert terminé",
                        "{0} fichier(s) traité(s)".format(burst_count),
                    )
                    _set_status("{0} — Prêt".format(BOX_NAME), processing=False)
                    burst_count = 0
                if _stop_event.is_set():
                    break
                continue

            # --- Stop signal ---
            if file is None:
                break

            log.info("Processing: %s (%d pending)", file.name, file_queue.qsize())

            if burst_count == 0 and not needs_refresh:
                _notify("Transfert en cours", file.name)
            _set_status("Transfert en cours...", processing=True)

            if not file.exists():
                log.warning("File gone before processing: %s", file)
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error("Aborting — persistent lock: %s", file.name)
                _notify("Erreur", "Fichier verrouillé : {0}".format(file.name))
                _log_medecin(
                    "Erreur : le fichier {0} est verrouillé et n'a pas pu être transféré.".format(
                        file.name
                    )
                )
                file_queue.task_done()
                continue

            # --- Wait for an open patient, up to PATIENT_WAIT_TIMEOUT seconds ---
            patient: Optional[dict] = None
            start_time = time.monotonic()
            first_log  = True

            while True:
                patient = get_active_patient()
                if patient:
                    break

                elapsed = time.monotonic() - start_time
                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    orphan_file(file)
                    _notify("Fichier orphelin", file.name)
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(
                        "No patient open — waiting (timeout in %d min).",
                        PATIENT_WAIT_TIMEOUT // 60,
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                "Patient: %s %s (code %s)",
                patient["nom"], patient["prenom"], patient["code"],
            )

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(
                    "Could not resolve folder for patient %s — orphaning.",
                    patient["code"],
                )
                orphan_file(file)
                _notify("Fichier orphelin", file.name)
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
                last_patient_code = patient["code"]
                burst_count      += 1
                log.debug("Insert OK — needs_refresh=True (deferred to burst end).")
                _log_medecin(
                    "Image transférée avec succès pour le patient {0} {1} ({2}).".format(
                        patient["nom"].upper(),
                        patient["prenom"],
                        description,
                    )
                )
            else:
                log.warning("Insert failed — refresh flag unchanged.")
                _notify("Erreur BD", "Insertion échouée — vérifier les logs")
                _log_medecin(
                    "Erreur : l'image {0} n'a pas pu être enregistrée pour le patient {1} {2}.".format(
                        file.name,
                        patient["nom"].upper(),
                        patient["prenom"],
                    )
                )

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                _notify(
                    "Transfert terminé",
                    "{0} fichier(s) traité(s)".format(burst_count),
                )
        _set_status("{0} — Arrêté".format(BOX_NAME))
        pythoncom.CoUninitialize()


 
# Catch-up sweep thread
#
# Runs every SWEEP_INTERVAL_SECONDS and re-enqueues any file in SOURCE_DIR
# that watchdog may have missed (network drop, late arrival, etc.).
# A shared set tracks files already enqueued to avoid double-processing.
 

_enqueued_files: Set[Path] = set()
_enqueued_lock  = threading.Lock()


def _sweep_source_dir(file_queue: queue.Queue) -> None:
    """
    Scan SOURCE_DIR for valid files that are not yet in the queue.
    Thread-safe via _enqueued_lock.
    """
    try:
        found = list(SOURCE_DIR.rglob("*"))
    except Exception as exc:
        log.warning("Sweep: could not list SOURCE_DIR: %s", exc)
        return

    for path in found:
        if not path.is_file():
            continue
        if path.suffix.lower() not in WATCHED_EXTENSIONS:
            continue
        with _enqueued_lock:
            if path in _enqueued_files:
                continue
            _enqueued_files.add(path)

        log.info("Sweep: re-enqueuing missed file: %s", path.name)
        _log_medecin(
            "Fichier détecté lors du balayage périodique (non capturé en temps réel) : {0}.".format(
                path.name
            )
        )
        file_queue.put(path)


def _run_sweep(file_queue: queue.Queue) -> None:
    """Background thread: periodic catch-up sweep."""
    log.info(
        "Sweep thread started — interval: %d s.",
        SWEEP_INTERVAL_SECONDS,
    )
    while not _stop_event.wait(timeout=SWEEP_INTERVAL_SECONDS):
        log.debug("Sweep: scanning SOURCE_DIR for missed files...")
        _sweep_source_dir(file_queue)
    log.info("Sweep thread stopped.")


 
# Watchdog producer
 

class ImageProducer(FileSystemEventHandler):
    """Enqueue newly created files. Also registers them in _enqueued_files
    so the sweep thread does not re-enqueue them unnecessarily."""

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event: object) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        with _enqueued_lock:
            if file in _enqueued_files:
                return
            _enqueued_files.add(file)
        log.info("Enqueued (watchdog): %s (queue: %d)", file.name, self._queue.qsize() + 1)
        self._queue.put(file)


 
# Background thread — PollingObserver + auto-reconnect
 

def _run_background(file_queue: queue.Queue) -> None:
    _RECONNECT_DELAY = 15

    def _start_observer() -> Observer:
        producer = ImageProducer(file_queue)
        obs = Observer()
        obs.schedule(producer, str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — watching: %s", SOURCE_DIR)
        return obs

    observer = _start_observer()
    _set_status("{0} — Prêt".format(BOX_NAME), processing=False)

    try:
        while not _stop_event.is_set():
            time.sleep(1)
            if not observer.is_alive():
                log.warning(
                    "Observer died (possible network drop) — waiting %d s before reconnect.",
                    _RECONNECT_DELAY,
                )
                _set_status("{0} — Reconnexion...".format(BOX_NAME))
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                time.sleep(_RECONNECT_DELAY)
                log.info("Restarting Observer after network recovery.")
                observer = _start_observer()
                _set_status("{0} — Prêt".format(BOX_NAME), processing=False)
    finally:
        observer.stop()
        observer.join()
        remaining = file_queue.qsize()
        if remaining:
            log.info("Waiting for %d remaining file(s)...", remaining)
            file_queue.join()
        log.info("Background thread stopped.")
        if _icon is not None:
            _icon.stop()


 
# Entry point
 

def main() -> None:
    global _icon, _mutex_handle

    # Single-instance guard.
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_Windows7_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

    # Block manual double-click restart when Studio Vision is running.
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
                "Pour relancer le routeur d'images, veuillez fermer "
                "complètement puis relancer Studio Vision.",
                "Routeur d'images",
                0x30,  # MB_ICONWARNING | MB_OK
            )
            sys.exit(0)

    prevent_sleep()

    log.info("Checking network share availability...")
    wait_for_network_share()

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 5 started")
    log.info("  Source        : %s", SOURCE_DIR)
    log.info("  Dest          : %s", DEST_PHOTOS)
    log.info("  PUBLIC.MDB    : %s", PUBLIC_MDB)
    log.info("  DOCUM.MDB     : %s", DOCUM_MDB)
    log.info("  Orphans       : %s", ORPHAN_DIR)
    log.info("  Timeout       : %d min", PATIENT_WAIT_TIMEOUT // 60)
    log.info("  Sweep every   : %d s", SWEEP_INTERVAL_SECONDS)
    log.info("  Extensions    : %s", ", ".join(sorted(WATCHED_EXTENSIONS)))
    log.info("  Tech log      : %s", _LOG_FILE_TECH)
    log.info("  Médecin log   : %s", _LOG_FILE_MEDECIN)

    _log_medecin("Routeur d'images démarré — surveillance active.")

    file_queue: queue.Queue = queue.Queue()

    worker_thread = threading.Thread(target=worker, args=(file_queue,), name="Worker")
    worker_thread.daemon = True
    worker_thread.start()

    bg_thread = threading.Thread(
        target=_run_background, args=(file_queue,), name="Background"
    )
    bg_thread.daemon = True
    bg_thread.start()

    sweep_thread = threading.Thread(
        target=_run_sweep, args=(file_queue,), name="Sweep"
    )
    sweep_thread.daemon = True
    sweep_thread.start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without system tray.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested via keyboard.")
        finally:
            _stop_event.set()
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Ouvrir les logs", _open_logs),
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
    _log_medecin("Routeur d'images arrêté.")
    log.info("Application stopped.")


if __name__ == "__main__":
    main()