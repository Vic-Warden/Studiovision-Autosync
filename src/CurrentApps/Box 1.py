"""
Medical imaging router — Version 6

Routes image files dropped by the acquisition system into the correct patient
folder on the network share, then inserts a record into the SFDoc subform of
the active Access form via win32com GUI automation.

Pipeline:
  PollingObserver → file_queue → Worker → move file → GUI insert (win32com)

Dependencies: watchdog, pywin32, pythoncom, pystray, Pillow, psutil
"""

import os
import pythoncom
import queue
import shutil
import subprocess
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
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

import win32api
import win32event
import winerror
import psutil


# Configuration
BOX_NAME = "Studiovision"

SOURCE_DIR   = Path(r"C:\Box 1")  # Acquisition drop folder
ORPHAN_DIR   = Path(r"C:\Users\Box-1\Desktop\Images_Oubliées")  # Destination for unmatched files
DEST_PHOTOS  = Path(r"\\studiovision\Studiov2000-OM\PHOTOS")  # Root of the network photo archive

STUDIO_VISION_CMD = [
    r"C:\Studiov2000-OM\svprog\msaccess.exe",
    "/runtime",
    r"C:\Studiov2000-OM\svprog\Ophprog.mde",
    "/wrkgrp",
    r"C:\Studiov2000-OM\config\system.mdw",
    "/User",
    "/Pwd",
    "/X",
    "demarrage",
]

WATCHED_EXTENSIONS: set[str] = {
    ".jpg", ".jpeg", ".jfif",
    ".png", ".bmp",
    ".tif", ".tiff",
    ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION: dict[str, str] = {
    ".jpg":  "OCT",
    ".jpeg": "OCT",
    ".jfif": "OCT",
    ".png":  "OCT",
    ".bmp":  "OCT",
    ".tif":  "OCT",
    ".tiff": "OCT",
    ".dcm":  "OCT",
    ".pdf":  "OCT",
    ".rtf":  "OCT",
    ".doc":  "OCT",
    ".docx": "OCT",
    ".odt":  "OCT",
}

FILE_LOCK_RETRY_DELAY:  int = 3
FILE_LOCK_MAX_ATTEMPTS: int = 15
PATIENT_POLL_INTERVAL:  int = 3
PATIENT_WAIT_TIMEOUT:   int = 900
CATCHUP_INTERVAL:       int = 120

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

SFDOC_SUBFORM_NAME = "SFDoc"
_AC_SUBFORM        = 112  # Access control type constant for subforms

_GUI_PRE_INSERT_DELAY = 0.3  # Seconds between file move and GUI insert
_UI_POST_INSERT_DELAY = 0.5  # Seconds between insert and Requery/MoveLast


# Logging
_LOG_DIR = Path(os.path.expanduser("~")) / "studiovision"
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


def _get_doctor_log_path() -> Path:
    return _LOG_DIR / f"doctor_report_{datetime.now().strftime('%Y-%m-%d')}.txt"


def log_doctor(message: str) -> None:
    """Appends a timestamped line to the daily doctor report."""
    timestamp = datetime.now().strftime("%H:%M")
    try:
        with open(_get_doctor_log_path(), "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as exc:
        log.warning(f"Could not write to doctor report: {exc}")


# Global state
_NETWORK_SHARE_POLL = 10
_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)

_icon: "pystray.Icon | None" = None
_status_text: str             = "Starting..."
_stop_event: threading.Event  = threading.Event()
_mutex_handle                 = None


# System tray helpers
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
        except Exception as exc:
            log.debug(f"Tray update failed: {exc}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as exc:
            log.debug(f"Notification failed: {exc}")


def _open_log_file(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as exc:
        log.warning(f"Could not open log file: {exc}")


def _open_doctor_log(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_get_doctor_log_path()))
    except Exception as exc:
        log.warning(f"Could not open doctor report: {exc}")


def _quit(icon, item) -> None:  # noqa: ARG001
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


# Network share
def wait_for_network_share() -> None:
    """Blocks until SOURCE_DIR is reachable."""
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


# Patient folder resolution
def build_patient_relative_path(patient_code: str, last_name: str, first_name: str) -> str:
    """
    Returns the relative folder path for a patient.
    Format: <first2digits>.000\\<code><last4>.<first3>
    Example: code=1758511228, DE GAULLE, CHARLES → "17.000\\1758511228dega.cha"
    """
    prefix   = patient_code[:2]
    clean    = str.maketrans("", "", " '-")
    last_4   = last_name.translate(clean).lower()[:4]
    first_3  = first_name.translate(clean).lower()[:3]
    return f"{prefix}.000\\{patient_code}{last_4}.{first_3}"


def resolve_patient_folder(patient: dict) -> Path | None:
    """Resolves and creates the absolute patient folder. Returns None on failure."""
    try:
        rel    = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
        folder = DEST_PHOTOS / rel
        folder.mkdir(parents=True, exist_ok=True)
        log.info(f"Patient folder resolved: {folder}")
        return folder
    except Exception as exc:
        log.error(f"Could not resolve/create patient folder: {exc}")
        return None


# Access COM — read active patient
def get_active_patient() -> dict | None:
    """Returns the active patient's code, last name, and first name from the Access form."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target: set[str] = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
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
        log.debug(f"COM error while reading patient: {exc}")
        return None


# Access COM — find SFDoc subform
def _find_sfdoc(form) -> object | None:
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
    Inserts a new record into the SFDoc subform via win32com.
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

        current = get_active_patient()
        if not current or current["code"] != patient["code"]:
            log.warning(
                f"GUI insert aborted: patient changed "
                f"(expected={patient['code']}, "
                f"current={current['code'] if current else 'none'})."
            )
            return False

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.error(f"Subform '{SFDOC_SUBFORM_NAME}' not found — GUI insertion aborted.")
            return False

        rs = sfdoc.Recordset
        rs.AddNew()

        def _set_field(name: str, value) -> None:
            try:
                rs.Fields(name).Value = value
            except Exception as exc:
                log.warning(f"Field '{name}' write failed: {exc}")

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
            f"GUI insert OK: patient={patient['code']} "
            f"desc='{description}' path='{relative_path}'"
        )

        time.sleep(_UI_POST_INSERT_DELAY)

        try:
            sfdoc.Requery()
            log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}'.")
        except Exception as exc:
            log.warning(f"Requery() failed: {exc}")
            try:
                sfdoc.Refresh()
                log.info(f"Fallback Refresh() on '{SFDOC_SUBFORM_NAME}'.")
            except Exception as exc2:
                log.warning(f"Fallback Refresh() also failed: {exc2}")

        try:
            sfdoc.Recordset.MoveLast()
        except Exception as exc:
            log.debug(f"MoveLast() failed: {exc}")

        return True

    except Exception as exc:
        log.error(f"GUI insertion failed: {exc}")
        return False


# File utilities
def wait_for_file(file: Path) -> bool:
    """Waits until the file is no longer locked. Returns False if timeout is reached."""
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
    """Moves source to dest_folder, resolving name conflicts with a timestamp suffix."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict resolved — renamed to {dest.name}")
    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} → {dest}")
        return dest
    except Exception as exc:
        log.error(f"Move failed: {exc}")
        return None


def orphan_file(file: Path) -> None:
    """Moves an unprocessable file to the orphan directory."""
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def prevent_sleep() -> None:
    """Prevents Windows from sleeping while the router is active."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as exc:
        log.warning(f"Could not set execution state: {exc}")


# Startup / catchup scan
def _scan_source_for_missed_files(file_queue: queue.Queue) -> None:
    """Enqueues files in SOURCE_DIR that were missed during downtime."""
    if not SOURCE_DIR.is_dir():
        log.warning(f"Catchup scan skipped: {SOURCE_DIR} not accessible.")
        return

    found: list[Path] = []
    try:
        for item in SOURCE_DIR.rglob("*"):
            if item.is_file() and item.suffix.lower() in WATCHED_EXTENSIONS:
                found.append(item)
    except Exception as exc:
        log.error(f"Error during catchup scan: {exc}")
        return

    if not found:
        log.info("Catchup scan: no pending files.")
        return

    log.info(f"Catchup scan: {len(found)} file(s) found — enqueuing.")
    for f in found:
        file_queue.put(f)
        log.info(f"  Catchup → enqueued: {f.name}")


def _catchup_loop(file_queue: queue.Queue) -> None:
    """Periodically re-scans SOURCE_DIR to catch files missed by the observer."""
    log.info(f"Catchup thread started (interval: {CATCHUP_INTERVAL}s).")
    while not _stop_event.is_set():
        for _ in range(CATCHUP_INTERVAL):
            if _stop_event.is_set():
                break
            time.sleep(1)
        if _stop_event.is_set():
            break
        log.debug("Catchup: scanning SOURCE_DIR...")
        _scan_source_for_missed_files(file_queue)
    log.info("Catchup thread stopped.")


# Worker
def worker(file_queue: queue.Queue) -> None:
    """
    Consumes files from the queue.
    For each file: wait for unlock → wait for open patient → move → GUI insert.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh:     bool       = False
    last_patient_code: str | None = None
    burst_count:       int        = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — all files in batch processed.")
                    _notify("Transfer complete", f"{burst_count} file(s) processed")
                    _set_status(f"{BOX_NAME} — Ready", processing=False)
                    needs_refresh     = False
                    last_patient_code = None
                    burst_count       = 0
                continue
            except Exception as exc:
                log.error(f"Queue error: {exc}")
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
                log.error(f"Aborting — persistent file lock: {file.name}")
                _notify("Error", f"File locked: {file.name}")
                file_queue.task_done()
                continue

            # Wait for an open patient in Access
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
                    log_doctor(
                        f"WARNING: File '{file.name}' orphaned "
                        f"(no patient open after {PATIENT_WAIT_TIMEOUT // 60} min)."
                    )
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(
                        f"No patient open — waiting "
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

            patient_folder = resolve_patient_folder(patient)
            if not patient_folder:
                log.error(f"Could not resolve folder for patient {patient['code']}. Orphaning.")
                orphan_file(file)
                _notify("Orphan file", file.name)
                log_doctor(
                    f"ERROR: Folder could not be created for '{file.name}' "
                    f"(patient {patient['nom']} {patient['prenom']}, "
                    f"code {patient['code']})."
                )
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            rel_path      = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
            relative_path = f"\\{rel_path}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            time.sleep(_GUI_PRE_INSERT_DELAY)

            if gui_insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.info(f"Record inserted: '{dest.name}' → {relative_path}")
                log_doctor(
                    f"SUCCESS: '{dest.name}' added to the record of "
                    f"{patient['nom']} {patient['prenom']} (code {patient['code']})."
                )
            else:
                log.error(
                    f"GUI insertion FAILED for '{dest.name}' "
                    f"(patient {patient['code']}). File is at: {dest}. Manual entry required."
                )
                _notify("Insert error", f"'{dest.name}' moved but not inserted — see log.")
                log_doctor(
                    f"ERROR: '{dest.name}' was moved to {dest} but could NOT "
                    f"be inserted for patient {patient['nom']} {patient['prenom']} "
                    f"(code {patient['code']}). Manual entry required."
                )

            file_queue.task_done()

    finally:
        _set_status(f"{BOX_NAME} — Stopped")
        pythoncom.CoUninitialize()
        log.info("Worker stopped.")


# Watchdog producer
class ImageProducer(FileSystemEventHandler):
    """Enqueues newly created files detected by the filesystem observer."""

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


# Background observer thread
def _run_background(file_queue: queue.Queue) -> None:
    """Starts and monitors the filesystem observer. Reconnects on network drops."""
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started.")
        return obs

    observer = _start_observer()
    _set_status(f"{BOX_NAME} — Ready", processing=False)

    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observer stopped (network drop?). Reconnecting...")
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
            log.info(f"Draining {remaining} remaining file(s) from queue...")
            file_queue.join()
        log.info("Background thread stopped.")
        if _icon is not None:
            _icon.stop()


# Studio Vision lifecycle
_SV_POLL_INTERVAL    = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT  = 30  # Seconds to wait for msaccess.exe to appear after launch


def _get_msaccess_pids() -> set[int]:
    """Returns the set of all running msaccess.exe PIDs."""
    pids: set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision() -> None:
    """
    Launches Studio Vision and tracks any new msaccess.exe instances.
    Handles the /runtime 2-stage startup relay correctly.
    Forces termination of newly created zombie processes on exit.
    """
    log.info(f"Launching Studio Vision: {' '.join(STUDIO_VISION_CMD)}")

    pids_before: set[int] = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        log.critical(f"Studio Vision executable not found. Shutting down.")
        _stop_event.set()
        return
    except Exception as exc:
        log.error(f"Could not launch Studio Vision: {exc}. Shutting down.")
        _stop_event.set()
        return

    log.info(f"Waiting up to {_SV_STARTUP_TIMEOUT}s for msaccess.exe to start...")
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
    tracked_pids: set[int] = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            
            current_pids = _get_msaccess_pids() - pids_before
            
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                log.debug(f"Studio Vision missing ({consecutive_empty}/{_EMPTY_THRESHOLD}).")
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    log.info("Studio Vision closed by user. Initiating shutdown.")
                    break
            else:
                consecutive_empty = 0
                
    except Exception as exc:
        log.error(f"Error while monitoring msaccess.exe: {exc}")
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    log.info(f"Force-killed zombie msaccess.exe (PID {pid}) to release COM locks.")
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
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Mutex")
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
                "To restart the image router, please close Studio Vision "
                "completely and relaunch it.",
                "Image Router",
                0x30,
            )
            sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Studio Vision Image Router — Version 6 ===")
    log.info(f"  Source dir     : {SOURCE_DIR}")
    log.info(f"  Dest photos    : {DEST_PHOTOS}")
    log.info(f"  Orphan dir     : {ORPHAN_DIR}")
    log.info(f"  Log file       : {_LOG_FILE}")
    log.info(f"  Patient timeout: {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Extensions     : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info(f"  Catchup every  : {CATCHUP_INTERVAL}s")
    log.info(f"  SFDoc subform  : {SFDOC_SUBFORM_NAME}")
    log.info(f"  Doctor log     : {_get_doctor_log_path()}")

    log_doctor(f"Image router started (version 6). Watching: {SOURCE_DIR}")

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(target=worker,          args=(file_queue,), name="Worker",               daemon=True).start()
    threading.Thread(target=_catchup_loop,   args=(file_queue,), name="Catchup",              daemon=True).start()
    threading.Thread(target=_run_background, args=(file_queue,), name="Background",           daemon=True).start()

    log.info("Startup scan — checking for pending files...")
    _scan_source_for_missed_files(file_queue)

    sv_thread = threading.Thread(target=_launch_studio_vision, name="StudioVisionLauncher", daemon=True)
    sv_thread.start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without system tray.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested via keyboard interrupt.")
        finally:
            _stop_event.set()
        log.info("Application stopped.")
        log_doctor("Image router stopped.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open technical log",  _open_log_file),
        pystray.MenuItem("Open doctor report",  _open_doctor_log),
        pystray.MenuItem("Quit",                _quit),
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
    log_doctor("Image router stopped.")


if __name__ == "__main__":
    main()
