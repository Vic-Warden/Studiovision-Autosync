"""
Medical imaging router — Version 6 (Windows 7 / Python 3.8.10)

Routes image files dropped by the acquisition system into the correct patient
folder on the network share, then inserts a record into the SFDoc subform of
the active Access form via win32com GUI automation.

Pipeline:
  PollingObserver → file_queue → Worker → move file → GUI insert (win32com)

A shared set (_enqueued_files) prevents double-processing between the watchdog
and the periodic sweep thread. Handles the /runtime 2-stage relay and
force-kills zombie COM processes on exit.

Dependencies: watchdog, pywin32, pythoncom, pystray, Pillow, psutil
"""

import ctypes
import logging
import logging.handlers
import os
import pythoncom
import queue
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, Set

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
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False


# Configuration
BOX_NAME = "Windows 7"

SOURCE_DIR  = Path(r"??")  # Acquisition drop folder
ORPHAN_DIR  = Path(r"??")  # Destination for unmatched files
DEST_PHOTOS = Path(r"??")  # Root of the network photo archive

STUDIO_VISION_CMD = [
    r"C:\Studiov2000-W7\svprog\msaccess.exe",
    "/runtime",
    r"C:\Studiov2000-W7\svprog\Ophprog.mde",
    "/wrkgrp",
    r"C:\Studiov2000-W7\config\system.mdw",
    "/User",
    "/Pwd",
    "/X",
    "demarrage",
]

WATCHED_EXTENSIONS: Set[str] = {
    ".jpg", ".jpeg", ".jfif",
    ".png", ".bmp",
    ".tif", ".tiff",
    ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION: Dict[str, str] = {
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

FILE_LOCK_RETRY_DELAY:  int = 3
FILE_LOCK_MAX_ATTEMPTS: int = 15
PATIENT_POLL_INTERVAL:  int = 3
PATIENT_WAIT_TIMEOUT:   int = 900
SWEEP_INTERVAL_SECONDS: int = 300  # 5 minutes

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"
SFDOC_SUBFORM_NAME  = "SFDoc"
_AC_SUBFORM         = 112  # Access ControlType constant for subform

_GUI_PRE_INSERT_DELAY = 0.3  # Seconds between file move and GUI insert
_UI_POST_INSERT_DELAY = 0.5  # Seconds between insert and Requery/MoveLast


# Logging — two timed-rotating handlers:
#   image_router.log      : full technical log (30-day rotation)
#   transferts_medecin.log: plain-French summary for end-users (90-day)
_LOG_DIR = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)

_LOG_FILE_TECH    = os.path.join(_LOG_DIR, "image_router.log")
_LOG_FILE_MEDECIN = os.path.join(_LOG_DIR, "transferts_medecin.log")

_tech_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_TECH, when="midnight", interval=1, backupCount=30, encoding="utf-8",
)
_tech_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)

_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(
    logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S")
)

logging.basicConfig(level=logging.INFO, handlers=[_tech_handler, _console_handler])
log = logging.getLogger("image_router")

_medecin_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_MEDECIN, when="midnight", interval=1, backupCount=90, encoding="utf-8",
)
_medecin_handler.setFormatter(logging.Formatter("%(message)s"))

_medecin_log = logging.getLogger("medecin")
_medecin_log.setLevel(logging.INFO)
_medecin_log.propagate = False  # Keep separate from the technical log
_medecin_log.addHandler(_medecin_handler)


def _log_medecin(msg: str) -> None:
    """Writes a timestamped plain-French line to the doctor-facing log."""
    _medecin_log.info("%s - %s", datetime.now().strftime("%H:%M"), msg)


# Global state
_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)

_icon:         Optional["pystray.Icon"] = None
_status_text:  str                      = "Démarrage..."
_stop_event:   threading.Event          = threading.Event()
_mutex_handle                           = None


# System tray helpers
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


def _open_tech_log(icon: object, item: object) -> None:
    try:
        os.startfile(_LOG_FILE_TECH)
    except Exception as exc:
        log.warning("Could not open technical log: %s", exc)


def _open_medecin_log(icon: object, item: object) -> None:
    try:
        os.startfile(_LOG_FILE_MEDECIN)
    except Exception as exc:
        log.warning("Could not open doctor log: %s", exc)


def _quit(icon: object, item: object) -> None:
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


# Network share
def wait_for_network_share() -> None:
    """Blocks until SOURCE_DIR is reachable."""
    source_str = str(SOURCE_DIR)
    is_local   = not (source_str.startswith("\\\\") or source_str.startswith("//")) \
                 and len(source_str) >= 2 and source_str[1] == ":"
    if is_local:
        return
    first = True
    while True:
        try:
            if SOURCE_DIR.is_dir():
                if not first:
                    log.info("Network share is now reachable: %s", SOURCE_DIR)
                return
        except Exception:
            pass
        log.warning("Network share not reachable, retrying in 10s: %s", SOURCE_DIR)
        first = False
        time.sleep(10)


# Patient folder resolution
def build_patient_relative_path(patient_code: str, last_name: str, first_name: str) -> str:
    """
    Returns the relative folder path using the Studio Vision naming convention.
    Format: <first2digits>.000\\<code><last3>.<first3>
    Example: code=1758511228, ABCDEF, DEFGH → "17.000\\1758511228abc.def"
    """
    prefix  = patient_code[:2]
    last_3  = last_name[:3].lower()
    first_3 = first_name[:3].lower()
    return "{0}.000\\{1}{2}.{3}".format(prefix, patient_code, last_3, first_3)


def resolve_patient_folder(patient: dict) -> Optional[Path]:
    """Resolves and creates the absolute patient folder on disk. Returns None on failure."""
    try:
        rel    = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
        folder = DEST_PHOTOS / rel
        folder.mkdir(parents=True, exist_ok=True)
        log.info("Patient folder resolved: %s", folder)
        return folder
    except Exception as exc:
        log.error("Could not resolve/create patient folder: %s", exc)
        return None


# Access COM — read active patient
def get_active_patient() -> Optional[dict]:
    """Returns the active patient's code, last name, and first name from the Access form."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target: Set[str] = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data:   dict     = {}

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


# Access COM — find SFDoc subform
def _find_sfdoc(form: object) -> Optional[object]:
    """Recursively searches the form's control tree for the SFDoc subform."""
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


# GUI insertion
def gui_insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """
    Inserts a new record into the SFDoc subform via win32com GUI automation.
    Calls AddNew(), fills fields, calls Update(), then Requery() + MoveLast().
    Returns True on success, False on any COM error.
    """
    if not WIN32_AVAILABLE:
        log.error("win32com not available — GUI insertion skipped.")
        return False

    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("GUI insert skipped: no active form in Access.")
            return False

        # Abort if the user switched to a different patient
        current = get_active_patient()
        if not current or current["code"] != patient["code"]:
            log.warning(
                "GUI insert aborted: patient changed (expected=%s, current=%s).",
                patient["code"], current["code"] if current else "none",
            )
            return False

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.error("Subform '%s' not found — GUI insertion aborted.", SFDOC_SUBFORM_NAME)
            return False

        rs = sfdoc.Recordset
        rs.AddNew()

        def _set_field(name: str, value: object) -> None:
            try:
                rs.Fields(name).Value = value
            except Exception as exc:
                log.warning("Field '%s' write failed: %s", name, exc)

        _set_field("code patient",  int(patient["code"]))
        _set_field("Date",          datetime.now())
        _set_field("DESCRIPTIONS",  description)
        _set_field("TEXTE",         relative_path)
        _set_field("Photo externe", relative_path)
        _set_field("TypeVW",        99)

        try:
            rs.Fields("NumDocExterne").Value = None
        except Exception:
            pass

        rs.Update()
        log.info(
            "GUI insert OK: patient=%s desc='%s' path='%s'",
            patient["code"], description, relative_path,
        )

        time.sleep(_UI_POST_INSERT_DELAY)

        try:
            sfdoc.Requery()
            log.info("Requery() on '%s' OK.", SFDOC_SUBFORM_NAME)
        except Exception as exc:
            log.warning("Requery() failed: %s — trying Refresh().", exc)
            try:
                sfdoc.Refresh()
                log.info("Fallback Refresh() on '%s' OK.", SFDOC_SUBFORM_NAME)
            except Exception as exc2:
                log.warning("Fallback Refresh() also failed: %s", exc2)

        try:
            sfdoc.Recordset.MoveLast()
        except Exception as exc:
            log.debug("MoveLast() failed: %s", exc)

        return True

    except Exception as exc:
        log.error("GUI insertion failed: %s", exc)
        return False


# File utilities
def wait_for_file(file: Path) -> bool:
    """Blocks until the file is readable. Returns False if max attempts are exceeded."""
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
    """Moves source to dest_folder, resolving name conflicts with a timestamp suffix."""
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
    """Moves an unprocessable file to the orphan directory."""
    log.warning("Orphaning: %s", file.name)
    move_file(file, ORPHAN_DIR, label="ORPHAN")
    _log_medecin(
        "Fichier non attribué (aucun patient ouvert) : {0} — déplacé dans le dossier orphelins.".format(
            file.name
        )
    )


def prevent_sleep() -> None:
    """Prevents Windows from sleeping while the router is active."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as exc:
        log.warning("Could not set execution state: %s", exc)


# Worker thread
def worker(file_queue: queue.Queue) -> None:
    """
    Consumes files from the queue.
    For each file: wait for unlock → wait for open patient → move file → GUI insert.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh:     bool          = False
    last_patient_code: Optional[str] = None
    burst_count:       int           = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — all files in batch processed.")
                    _notify("Transfert terminé", "{0} fichier(s) traité(s)".format(burst_count))
                    _set_status("{0} — Prêt".format(BOX_NAME), processing=False)
                    needs_refresh     = False
                    last_patient_code = None
                    burst_count       = 0
                if _stop_event.is_set():
                    break
                continue

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
                    "Erreur : le fichier {0} est verrouillé et n'a pas pu être transféré.".format(file.name)
                )
                file_queue.task_done()
                continue

            # Wait for an open patient in Access
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
                    log.info("No patient open — waiting (timeout in %d min).", PATIENT_WAIT_TIMEOUT // 60)
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info("Patient: %s %s (code %s)", patient["nom"], patient["prenom"], patient["code"])

            patient_folder = resolve_patient_folder(patient)
            if not patient_folder:
                log.error("Could not resolve folder for patient %s — orphaning.", patient["code"])
                orphan_file(file)
                _notify("Fichier orphelin", file.name)
                _log_medecin(
                    "Erreur : impossible de créer le dossier pour {0} {1} "
                    "(code {2}) — fichier orphelin.".format(
                        patient["nom"].upper(), patient["prenom"], patient["code"]
                    )
                )
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            rel_path      = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
            relative_path = "\\{0}\\{1}".format(rel_path, dest.name)
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            time.sleep(_GUI_PRE_INSERT_DELAY)

            if gui_insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.info("Record inserted: '%s' -> %s", dest.name, relative_path)
                _log_medecin(
                    "Image transférée avec succès pour le patient {0} {1} ({2}).".format(
                        patient["nom"].upper(), patient["prenom"], description,
                    )
                )
            else:
                log.error(
                    "GUI insertion FAILED for '%s' (patient %s). File is at: %s. Manual entry required.",
                    dest.name, patient["code"], dest,
                )
                _notify("Erreur insertion", "'{0}' déplacé mais non inséré — voir logs.".format(dest.name))
                _log_medecin(
                    "Erreur : l'image {0} a été déplacée vers {1} mais n'a PAS pu "
                    "être insérée pour le patient {2} {3} (code {4}). "
                    "Saisie manuelle requise.".format(
                        dest.name, dest,
                        patient["nom"].upper(), patient["prenom"], patient["code"],
                    )
                )

            file_queue.task_done()

    finally:
        _set_status("{0} — Arrêté".format(BOX_NAME))
        pythoncom.CoUninitialize()
        log.info("Worker stopped.")


# Catch-up sweep — shared enqueue registry to prevent double-processing
_enqueued_files: Set[Path] = set()
_enqueued_lock  = threading.Lock()


def _sweep_source_dir(file_queue: queue.Queue) -> None:
    """Scans SOURCE_DIR and enqueues any valid file not already tracked."""
    try:
        found = list(SOURCE_DIR.rglob("*"))
    except Exception as exc:
        log.warning("Sweep: could not list SOURCE_DIR: %s", exc)
        return

    for path in found:
        if not path.is_file() or path.suffix.lower() not in WATCHED_EXTENSIONS:
            continue
        with _enqueued_lock:
            if path in _enqueued_files:
                continue
            _enqueued_files.add(path)

        log.info("Sweep: re-enqueuing missed file: %s", path.name)
        _log_medecin(
            "Fichier détecté lors du balayage périodique "
            "(non capturé en temps réel) : {0}.".format(path.name)
        )
        file_queue.put(path)


def _run_sweep(file_queue: queue.Queue) -> None:
    """Background thread: periodic catch-up sweep."""
    log.info("Sweep thread started — interval: %d s.", SWEEP_INTERVAL_SECONDS)
    while not _stop_event.wait(timeout=SWEEP_INTERVAL_SECONDS):
        log.debug("Sweep: scanning SOURCE_DIR for missed files...")
        _sweep_source_dir(file_queue)
    log.info("Sweep thread stopped.")


# Watchdog producer
class ImageProducer(FileSystemEventHandler):
    """Enqueues newly created files detected by the filesystem observer.
    Also registers them in _enqueued_files to avoid redundant sweep re-enqueuing."""

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


# Background observer thread
def _run_background(file_queue: queue.Queue) -> None:
    """Starts and monitors the filesystem observer. Reconnects on network drops."""
    _RECONNECT_DELAY = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
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
                    "Observer died (possible network drop) — waiting %ds before reconnect.",
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


# Studio Vision lifecycle
_SV_POLL_INTERVAL   = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT = 30  # Seconds to wait for msaccess.exe to appear after launch


def _get_msaccess_pids() -> Set[int]:
    """Returns the set of all running msaccess.exe PIDs."""
    pids: Set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision() -> None:
    """
    Launches Studio Vision and monitors all msaccess.exe processes.
    Handles the /runtime 2-stage relay: the launcher process exits ~2 s after
    start and spawns the real worker process under a new PID. We track the full
    set of new PIDs rather than a single one.
    Force-kills any tracked processes on exit to release COM locks.
    """
    log.info("Launching Studio Vision: %s", " ".join(STUDIO_VISION_CMD))

    pids_before: Set[int] = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        log.critical("Studio Vision executable not found. Shutting down.")
        _stop_event.set()
        return
    except Exception as exc:
        log.error("Could not launch Studio Vision: %s. Shutting down.", exc)
        _stop_event.set()
        return

    log.info("Waiting up to %ds for msaccess.exe to start...", _SV_STARTUP_TIMEOUT)
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            log.info("Studio Vision is starting...")
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            log.error("msaccess.exe did not appear. Shutting down.")
            _stop_event.set()
        return

    consecutive_empty = 0
    _EMPTY_THRESHOLD  = 2
    tracked_pids: Set[int] = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            current_pids = _get_msaccess_pids() - pids_before
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                log.debug("Studio Vision missing (%d/%d).", consecutive_empty, _EMPTY_THRESHOLD)
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    log.info("Studio Vision closed by user. Initiating shutdown.")
                    break
            else:
                consecutive_empty = 0

    except Exception as exc:
        log.error("Error while monitoring msaccess.exe: %s", exc)
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    log.info("Force-killed zombie msaccess.exe (PID %d) to release COM locks.", pid)
            except Exception:
                pass

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass


# Entry point
def main() -> None:
    global _icon, _mutex_handle

    # Single-instance guard
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_Windows7_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        router_alive = any(
            (p.info["name"] or "").lower() in ("python.exe", "pythonw.exe")
            and p.info["pid"] != os.getpid()
            for p in psutil.process_iter(["pid", "name"])
        )
        if router_alive:
            log.warning("Another instance is already running. Exiting.")
            sys.exit(0)
        else:
            log.warning("Stale mutex detected (previous crash). Continuing.")

    # Prevent manual restart while Studio Vision is running
    try:
        parent_name = psutil.Process(os.getpid()).parent().name().lower()
    except Exception:
        parent_name = ""

    if parent_name == "explorer.exe":
        sv_running = any(
            (p.info["name"] or "").lower() == "msaccess.exe"
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

    if not SOURCE_DIR.exists():
        log.critical("Source folder not found: %s", SOURCE_DIR)
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Studio Vision Image Router — Version 6 (Windows 7) ===")
    log.info("  Source dir     : %s", SOURCE_DIR)
    log.info("  Dest photos    : %s", DEST_PHOTOS)
    log.info("  Orphan dir     : %s", ORPHAN_DIR)
    log.info("  Tech log       : %s", _LOG_FILE_TECH)
    log.info("  Médecin log    : %s", _LOG_FILE_MEDECIN)
    log.info("  Patient timeout: %d min", PATIENT_WAIT_TIMEOUT // 60)
    log.info("  Extensions     : %s", ", ".join(sorted(WATCHED_EXTENSIONS)))
    log.info("  Sweep every    : %d s", SWEEP_INTERVAL_SECONDS)
    log.info("  SFDoc subform  : %s", SFDOC_SUBFORM_NAME)

    _log_medecin("Routeur d'images démarré (version 6) — surveillance active.")

    file_queue: queue.Queue = queue.Queue()

    log.info("Startup scan — checking for pending files...")
    _sweep_source_dir(file_queue)

    threading.Thread(target=worker,          args=(file_queue,), name="Worker",     daemon=True).start()
    threading.Thread(target=_run_background, args=(file_queue,), name="Background", daemon=True).start()
    threading.Thread(target=_run_sweep,      args=(file_queue,), name="Sweep",      daemon=True).start()

    sv_thread = threading.Thread(target=_launch_studio_vision, name="StudioVisionLauncher", daemon=True)
    sv_thread.start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without system tray.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested via keyboard.")
        finally:
            _stop_event.set()
        _log_medecin("Routeur d'images arrêté.")
        log.info("Application stopped.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Ouvrir le log technique", _open_tech_log),
        pystray.MenuItem("Ouvrir le log médecin",   _open_medecin_log),
        pystray.MenuItem("Quitter",                 _quit),
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