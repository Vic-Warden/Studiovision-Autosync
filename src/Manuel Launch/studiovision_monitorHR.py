"""
Studio Vision HR Monitor — Version 6.0

Routes incoming imaging files to the correct patient folder and inserts
the record directly into the active Access form via win32com GUI automation.
The pyodbc/SQL layer has been completely removed.

Pipeline:
  PollingObserver → file_queue → Worker → shutil.move + COM insert + UI refresh

Handles the /runtime 2-stage relay: _launch_studio_vision() tracks msaccess.exe
across the launcher-then-real-process handoff and force-kills zombie COM locks on exit.

Logging:
  Technical log : ~/studiovision/image_router_HR.log
  Doctor report : ~/studiovision/rapport_medecin_YYYY-MM-DD.txt (French)

Dependencies: watchdog, pywin32, pythoncom, pystray, Pillow, psutil
"""

import ctypes
import logging
import os
import pythoncom
import queue
import shutil
import subprocess
import sys
import threading
import time
import win32api
import win32com.client
import win32event
import winerror
import psutil
from datetime import datetime
from pathlib import Path
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False


# Configuration
BOX_NAME = "Studiovision HR"

SOURCE_DIR    = Path(r"C:\Users\Box-6\Desktop\HR")                  # Acquisition drop folder
ORPHAN_DIR    = Path(r"C:\Users\Box-6\Desktop\Images_Oubliées")     # Unassignable file quarantine
DEST_PHOTOS   = Path(r"\\studiovision\Studiov2000\PHOTOS")           # Root of the patient photo tree
TRIAGE_SCRIPT = r"C:\INFO-BOX-CV\Auto Récup\studiovision_export.py"  # Dispatcher script

STUDIO_VISION_CMD = [
    r"C:\Studiov2000\Svprog\MSACCESS.EXE",
    "/runtime", r"C:\Studiov2000\svprog\Ophprog.mde",
    "/wrkgrp",  r"C:\Studiov2000\config\system.mdw",
    "/User", "/Pwd", "/X", "demarrage",
]

WATCHED_EXTENSIONS = {
    ".jpg", ".jpeg", ".jfif", ".png", ".bmp",
    ".tif", ".tiff", ".dcm", ".pdf",
    ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION = {ext: "Champ visuel" for ext in WATCHED_EXTENSIONS}

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

SFDOC_SUBFORM_NAME = "SFDoc"
_AC_SUBFORM        = 112  # Access ControlType constant for subform

FILE_LOCK_RETRY_DELAY  = 3    # Seconds between lock-check attempts
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3    # Seconds between "is a patient open?" polls
PATIENT_WAIT_TIMEOUT   = 900  # Seconds before orphaning (15 min)
CATCHUP_INTERVAL       = 120  # Seconds between periodic source-dir scans

_SV_POLL_INTERVAL   = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT = 30  # Seconds to wait for msaccess.exe to appear after launch


# Logging
_LOG_DIR = Path(os.path.expanduser("~")) / "studiovision"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / "image_router_HR.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router_HR")


def _get_medecin_log_path() -> Path:
    return _LOG_DIR / f"rapport_medecin_{datetime.now().strftime('%Y-%m-%d')}.txt"


def log_medecin(message: str) -> None:
    """Appends a timestamped line to the daily doctor report (French)."""
    try:
        with open(_get_medecin_log_path(), "a", encoding="utf-8") as fh:
            fh.write(f"[{datetime.now().strftime('%Hh%M')}] {message}\n")
    except Exception as exc:
        log.warning(f"Could not write to doctor report: {exc}")


# Global state
_icon:         "pystray.Icon | None" = None
_status_text:  str                   = "Starting..."
_stop_event:   threading.Event       = threading.Event()
_mutex_handle                        = None
_sv_access_app = None  # COM Access.Application bound to the PID we launched

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)


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
            log.debug(f"Tray update failed: {exc}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message or title, title)
        except Exception as exc:
            log.debug(f"Notification failed: {exc}")


def _open_logs(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as exc:
        log.warning(f"Could not open log file: {exc}")


def _open_medecin_log(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_get_medecin_log_path()))
    except Exception as exc:
        log.warning(f"Could not open doctor report: {exc}")


def _quit(icon, item) -> None:  # noqa: ARG001
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


# System utilities
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


# Patient folder resolution
def _resolve_patient_folder(code: str, nom: str, prenom: str) -> "Path | None":
    """
    Derives the patient photo folder using the Studio Vision naming convention.
    Format: DEST_PHOTOS / <first2digits>.000 / <code><nom3>.<prenom3>
    Returns the resolved Path if it exists on disk, otherwise None.
    """
    code_str     = str(code).strip()
    nom_clean    = nom.strip().lower()
    prenom_clean = prenom.strip().lower()

    if len(code_str) < 2:
        log.error(f"Patient code too short to resolve folder: '{code_str}'")
        return None

    folder = DEST_PHOTOS / f"{code_str[:2]}.000" / f"{code_str}{nom_clean[:3]}.{prenom_clean[:3]}"

    if not folder.is_dir():
        log.error(f"Resolved folder does not exist: {folder} (patient {code_str} / {nom.upper()} {prenom})")
        return None

    log.info(f"Patient folder resolved: {folder}")
    return folder


# Access COM — read active patient
def _get_access_app():
    """
    Returns the Access.Application COM object bound to this instance.
    Scans the ROT to avoid returning the wrong instance when multiple
    msaccess.exe processes are running. Falls back to GetActiveObject.
    """
    global _sv_access_app
    if _sv_access_app is not None:
        try:
            _ = _sv_access_app.Version
            return _sv_access_app
        except Exception:
            _sv_access_app = None

    try:
        import win32com.client
        ctx  = pythoncom.CreateBindCtx(0)
        rot  = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(10)
            if not monikers:
                break
            for mk in monikers:
                try:
                    name = mk.GetDisplayName(ctx, None)
                    if "Access" not in name:
                        continue
                    obj = rot.GetObject(mk)
                    app = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                    return app
                except Exception:
                    pass
    except Exception:
        pass

    # Last resort.
    try:        return win32com.client.GetActiveObject("Access.Application")
    except Exception:
        return None

def get_active_patient() -> "dict | None":
    """Returns the active patient's code, last name, and first name from the Access form."""
    try:
        access = _get_access_app()
        if access is None:
            return None
        form = access.Screen.ActiveForm
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
            "code":   str(data[ACCESS_FIELD_CODE]).strip(),
            "nom":    str(data[ACCESS_FIELD_NOM]).strip(),
            "prenom": str(data[ACCESS_FIELD_PRENOM]).strip(),
        }
    except Exception as exc:
        log.debug(f"COM get_active_patient error: {exc}")
        return None

# Access COM — find SFDoc subform
def _find_sfdoc(form):
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
def _insert_via_com(patient: dict, relative_path: str, description: str) -> bool:
    """
    Inserts a new record into the SFDoc subform via COM automation.
    Returns True on success, False on any COM error.
    """
    try:
        access = _get_access_app()
        if access is None:
            log.error("COM insert failed: could not obtain Access.Application.")
            return False

        form = access.Screen.ActiveForm
        if form is None:
            log.error("COM insert failed: no active Access form.")
            return False

        current = get_active_patient()
        if not current or current["code"] != patient["code"]:
            log.warning(
                f"COM insert aborted: patient changed "
                f"(expected={patient['code']}, "
                f"current={current['code'] if current else 'none'})."
            )
            return False

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.error(f"COM insert failed: subform '{SFDOC_SUBFORM_NAME}' not found.")
            return False

        rs = sfdoc.Recordset
        rs.AddNew()

        def _set(name, value):
            try:
                rs.Fields(name).Value = value
            except Exception as e:
                log.warning(f"  Field '{name}' write failed: {e}")

        from datetime import datetime as _dt
        _set("code patient",  int(patient["code"]))
        _set("Date",          _dt.now())
        _set("DESCRIPTIONS",  description)
        _set("TEXTE",         relative_path)
        _set("Photo externe", relative_path)
        _set("TypeVW",        99)
        try:
            rs.Fields("NumDocExterne").Value = None
        except Exception:
            pass

        rs.Update()
        log.info(
            f"COM insert OK: patient={patient['code']} "
            f"path='{relative_path}' desc='{description}'"
        )

        time.sleep(0.5)

        requery_ok = False
        for attempt in range(1, 4):
            try:
                sfdoc.Requery()
                log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}' (attempt {attempt}).")
                requery_ok = True
                break
            except Exception as exc_rq:
                log.warning(f"Requery() attempt {attempt}/3 failed: {exc_rq}")
                if attempt < 3:
                    time.sleep(0.5)

        if not requery_ok:
            log.warning(f"All Requery() attempts failed; falling back to Refresh().")
            try:
                sfdoc.Refresh()
                log.info("Fallback Refresh() OK.")
            except Exception as exc_ref:
                log.warning(f"Fallback Refresh() also failed: {exc_ref}")

        try:
            sfdoc.Recordset.MoveLast()
        except Exception:
            pass

        return True

    except Exception as exc:
        log.error(f"COM insert/refresh failed: {exc}")
        return False

# File utilities
def wait_for_file(file: Path) -> bool:
    """Blocks until the file is readable. Returns False if max attempts are exceeded."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retrying...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error(f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {file}")
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> "Path | None":
    """Moves source to dest_folder, resolving name conflicts with a timestamp suffix."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict — renamed to {dest.name}")
    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} -> {dest}")
        return dest
    except Exception as exc:
        log.error(f"Move failed: {exc}")
        return None


def orphan_file(file: Path) -> None:
    """Moves an unprocessable file to the orphan directory."""
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


# Startup scan and periodic catchup
def _scan_source_for_missed_files(file_queue: queue.Queue) -> None:
    """Enqueues files present in SOURCE_DIR that were missed during downtime."""
    if not SOURCE_DIR.is_dir():
        log.warning(f"Catchup scan skipped: {SOURCE_DIR} not accessible.")
        return
    found: list[Path] = []
    try:
        for item in SOURCE_DIR.iterdir():
            if item.is_file() and item.suffix.lower() in WATCHED_EXTENSIONS:
                found.append(item)
    except Exception as exc:
        log.error(f"Error during catchup scan: {exc}")
        return

    if not found:
        log.info("Catchup scan: no pending files.")
        return

    log.info(f"Catchup scan: {len(found)} file(s) found, enqueuing.")
    for f in found:
        file_queue.put(f)
        log.info(f"  Catchup -> enqueued: {f.name}")


def _catchup_loop(file_queue: queue.Queue) -> None:
    """Periodically re-scans SOURCE_DIR for files missed by the observer."""
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


# Worker thread
def worker(file_queue: queue.Queue) -> None:
    """
    Consumes files from the queue.
    For each file: wait for unlock → wait for open patient → resolve folder
    → move file → COM insert.
    """
    pythoncom.CoInitialize()
    log.info("Worker started.")

    burst_count: int = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if burst_count:
                    _notify("Transfer complete", f"{burst_count} file(s) processed")
                    _set_status(f"{BOX_NAME} — Ready", processing=False)
                    burst_count = 0
                continue
            except Exception as exc:
                log.error(f"Queue error: {exc}")
                continue

            log.info(f"Processing: {file.name} ({file_queue.qsize()} pending)")

            if burst_count == 0:
                _notify("Transfer in progress", file.name)
            _set_status("Transfer in progress...", processing=True)

            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Aborting — persistent lock: {file.name}")
                _notify("Error", f"File locked: {file.name}")
                file_queue.task_done()
                continue

            # Wait for an open patient record in Access
            patient    = None
            start_time = time.monotonic()
            first_log  = True

            while True:
                patient = get_active_patient()
                if patient:
                    break

                elapsed = time.monotonic() - start_time
                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    try:
                        orphan_file(file)
                        _notify("Orphan file", file.name)
                        log_medecin(
                            f"AVERTISSEMENT : Fichier '{file.name}' orphelin "
                            f"(aucun patient ouvert après {PATIENT_WAIT_TIMEOUT // 60} min)."
                        )
                    finally:
                        file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(f"No patient open — waiting (timeout in {PATIENT_WAIT_TIMEOUT // 60} min)")
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(f"Patient: {patient['nom']} {patient['prenom']} (code {patient['code']})")

            patient_folder = _resolve_patient_folder(patient["code"], patient["nom"], patient["prenom"])
            if not patient_folder:
                log.error(f"Could not resolve folder for patient {patient['code']}. Orphaning.")
                orphan_file(file)
                _notify("Orphan file", file.name)
                log_medecin(
                    f"ERREUR : Dossier introuvable pour '{file.name}' "
                    f"(patient {patient['nom']} {patient['prenom']}, Code: {patient['code']})."
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

            if _insert_via_com(patient, relative_path, description):
                burst_count += 1
                log_medecin(
                    f"SUCCÈS : Photo '{dest.name}' ajoutée au dossier de "
                    f"{patient['nom']} {patient['prenom']} (Code: {patient['code']})."
                )
            else:
                log.error(
                    f"COM insert failed for '{dest.name}' — file moved but record not created. "
                    "Manual insertion may be required."
                )
                _notify("COM Error", f"Insert failed for {dest.name}")
                log_medecin(
                    f"ERREUR COM : Photo '{dest.name}' déplacée mais NON insérée "
                    f"dans le dossier de {patient['nom']} {patient['prenom']} "
                    f"(Code: {patient['code']}). Insertion manuelle requise."
                )

            file_queue.task_done()

    finally:
        _set_status(f"{BOX_NAME} — Stopped")
        pythoncom.CoUninitialize()


# Watchdog producer
class ImageProducer(FileSystemEventHandler):
    """
    Watchdog event handler for the HR staging folder.
    A 5-second seen-set prevents duplicate enqueuing across created/moved/modified events.
    """

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue     = file_queue
        self._seen:     set[str] = set()
        self._seen_lock = threading.Lock()

    def _enqueue(self, path: Path, reason: str) -> None:
        key = str(path)
        with self._seen_lock:
            if key in self._seen:
                log.debug(f"Duplicate event suppressed ({reason}): {path.name}")
                return
            self._seen.add(key)

        def _clear_key():
            time.sleep(5)
            with self._seen_lock:
                self._seen.discard(key)

        threading.Thread(target=_clear_key, daemon=True).start()
        log.info(f"Enqueued ({reason}): {path.name} (queue size: {self._queue.qsize() + 1})")
        self._queue.put(path)

    def on_created(self, event) -> None:
        if not event.is_directory:
            path = Path(event.src_path)
            if path.suffix.lower() in WATCHED_EXTENSIONS:
                self._enqueue(path, "created")

    def on_moved(self, event) -> None:
        if not event.is_directory:
            path = Path(event.dest_path)
            if path.suffix.lower() in WATCHED_EXTENSIONS:
                self._enqueue(path, "moved")

    def on_modified(self, event) -> None:
        if not event.is_directory:
            path = Path(event.src_path)
            if path.suffix.lower() in WATCHED_EXTENSIONS and path.exists():
                self._enqueue(path, "modified")


# Background observer thread
def _run_background(file_queue: queue.Queue) -> None:
    """Starts and monitors the filesystem observer. Reconnects on failures."""
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=False)
        obs.start()
        log.info("Observer started.")
        return obs

    observer = _start_observer()
    _set_status(f"{BOX_NAME} — Ready", processing=False)

    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observer stopped. Attempting reconnect...")
                _set_status(f"{BOX_NAME} — Reconnecting...", processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
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


# Studio Vision lifecycle
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
    Launches Studio Vision HR and monitors all msaccess.exe processes.
    Handles the /runtime 2-stage relay by diffing PIDs before/after launch.
    Force-kills tracked processes on exit to release COM locks.
    """
    log.info(f"Launching Studio Vision: {' '.join(STUDIO_VISION_CMD)}")
    pids_before: set[int] = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        log.critical("Studio Vision executable not found. Shutting down.")
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


    # Bind COM Access.Application to the correct instance.
    global _sv_access_app
    try:
        import win32com.client as _wcc
        import pythoncom as _pycom
        _new_pids = sorted(_get_msaccess_pids() - pids_before)
        ctx  = _pycom.CreateBindCtx(0)
        rot  = _pycom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        _bound = False
        while not _bound:
            mks = enum.Next(10)
            if not mks:
                break
            for mk in mks:
                try:
                    name = mk.GetDisplayName(ctx, None)
                    if "Access" not in name:
                        continue
                    obj = rot.GetObject(mk)
                    app = _wcc.Dispatch(obj.QueryInterface(_pycom.IID_IDispatch))
                    _sv_access_app = app
                    log.info(f"COM Access.Application bound to this instance (PIDs={_new_pids}).")
                    _bound = True
                    break
                except Exception:
                    pass
        if not _bound:
            log.warning("Could not bind COM app to specific PID; will use GetActiveObject fallback.")
    except Exception as e:
        log.warning(f"COM binding failed: {e}")

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
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_HR_V6_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

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
                "To restart the image router, please fully close and relaunch Studio Vision.",
                "Image Router HR",
                0x30,
            )
            sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Studio Vision Image Router — Version 6.0 (HR) ===")
    log.info(f"  Source      : {SOURCE_DIR}")
    log.info(f"  Dest photos : {DEST_PHOTOS}")
    log.info(f"  Orphans     : {ORPHAN_DIR}")
    log.info(f"  Log file    : {_LOG_FILE}")
    log.info(f"  Timeout     : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Extensions  : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info(f"  Catchup     : every {CATCHUP_INTERVAL}s")
    log.info(f"  Doctor log  : {_get_medecin_log_path()}")

    log_medecin(f"Démarrage du routeur d'images (version 6.0 HR). Surveillance de : {SOURCE_DIR}")

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(target=worker,        args=(file_queue,), name="Worker",   daemon=True).start()
    threading.Thread(target=_catchup_loop, args=(file_queue,), name="Catchup",  daemon=True).start()

    log.info("Startup scan — checking for pending files...")
    _scan_source_for_missed_files(file_queue)

    threading.Thread(target=_run_background, args=(file_queue,), name="Background", daemon=True).start()

    if not TRAY_AVAILABLE:
        log.info("Headless mode — launching dispatcher and Studio Vision HR...")
        try:
            subprocess.Popen(["pythonw.exe", TRIAGE_SCRIPT])
            log.info("Dispatcher launched.")
        except Exception as exc:
            log.error(f"Could not launch dispatcher: {exc}")

        _launch_studio_vision()
        log.info("Studio Vision HR closed — HR monitor stopped.")
        log_medecin("Arrêt du routeur d'images (version 6.0 HR).")
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open technical log", _open_logs),
        pystray.MenuItem("Open doctor report", _open_medecin_log),
        pystray.MenuItem("Quit",               _quit),
    )

    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=_make_icon(_COLOR_READY),
        title=BOX_NAME,
        menu=menu,
    )

    try:
        subprocess.Popen(["pythonw.exe", TRIAGE_SCRIPT])
        log.info("Dispatcher launched.")
    except Exception as exc:
        log.error(f"Could not launch dispatcher: {exc}")

    threading.Thread(target=_launch_studio_vision, name="SVLifecycle", daemon=True).start()

    log.info("System tray icon started.")
    _icon.run()

    _stop_event.set()
    log.info("Application stopped.")
    log_medecin("Arrêt du routeur d'images (version 6.0 HR).")


if __name__ == "__main__":
    main()