"""
Medical imaging router — Version 6 Multi-Instance

Routes image files dropped by the acquisition system into the correct patient
folder on the network share, then inserts a record into the SFDoc subform of
the active Access form via win32com GUI automation.

Multi-instance: supports multiple simultaneous StudioVision instances.
Patient detection uses the ROT (Running Object Table) targeting only the
fPATIENTS form — no database queries.

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
from dataclasses import dataclass
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
import win32gui
import win32process
import win32pipe
import win32file
import pywintypes


# Configuration
BOX_NAME = "Studiovision"

SOURCE_DIR  = Path(r"C:\Users\Box-6\Desktop\Export CV")       # Acquisition drop folder
ORPHAN_DIR  = Path(r"C:\Users\Box-6\Desktop\Images_Oubliées")  # Unmatched file quarantine

# Each entry represents one StudioVision instance on this machine.
# mde_path  : local .mde path as it appears in the Windows ROT
# dest_photos: UNC root of the PHOTOS folder for this instance

@dataclass
class Instance:
    name:              str
    alias:             str    # shortcut argument (e.g. "HR", "OM")
    mde_path:          Path   # local .mde path (ROT)
    dest_photos:       Path   # network PHOTOS root
    studio_vision_cmd: list   # launch command


INSTANCES: list[Instance] = [
    Instance(
        name        = "Megret",
        alias       = "OM",
        mde_path    = Path(r"C:\Studiov2000-OM\svprog\Ophprog.mde"),
        dest_photos = Path(r"\\studiovision\Studiov2000-OM\PHOTOS"),
        studio_vision_cmd = [
            r"C:\Studiov2000-OM\svprog\msaccess.exe",
            "/runtime",
            r"C:\Studiov2000-OM\svprog\Ophprog.mde",
            "/wrkgrp", r"C:\Studiov2000-OM\config\system.mdw",
            "/User", "/Pwd", "/X", "demarrage",
        ],
    ),
    Instance(
        name        = "Romoli",
        alias       = "HR",
        mde_path    = Path(r"C:\Studiov2000\svprog\Ophprog.mde"),
        dest_photos = Path(r"\\studiovision\Studiov2000\PHOTOS"),
        studio_vision_cmd = [
            r"C:\Studiov2000\svprog\msaccess.exe",
            "/runtime",
            r"C:\Studiov2000\svprog\Ophprog.mde",
            "/wrkgrp", r"C:\Studiov2000\config\system.mdw",
            "/User", "/Pwd", "/X", "demarrage",
        ],
    ),
]

# Quick alias lookup
_INSTANCE_BY_ALIAS: dict[str, Instance] = {
    inst.alias.upper(): inst for inst in INSTANCES
}

WATCHED_EXTENSIONS: set[str] = {
    ".jpg", ".jpeg", ".jfif",
    ".png", ".bmp",
    ".tif", ".tiff",
    ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION: dict[str, str] = {
    ".jpg":  "Champ Visuel",
    ".jpeg": "Champ Visuel",
    ".jfif": "Champ Visuel",
    ".png":  "Champ Visuel",
    ".bmp":  "Champ Visuel",
    ".tif":  "Champ Visuel",
    ".tiff": "Champ Visuel",
    ".dcm":  "Champ Visuel",
    ".pdf":  "Champ Visuel",
    ".rtf":  "Champ Visuel",
    ".doc":  "Champ Visuel",
    ".docx": "Champ Visuel",
    ".odt":  "Champ Visuel",
}

FILE_LOCK_RETRY_DELAY:  int = 3
FILE_LOCK_MAX_ATTEMPTS: int = 15
PATIENT_POLL_INTERVAL:  int = 3
PATIENT_WAIT_TIMEOUT:   int = 900
CATCHUP_INTERVAL:       int = 120

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

PATIENT_FORM_NAME  = "fPATIENTS"
SFDOC_SUBFORM_NAME = "SFDoc"
_AC_SUBFORM        = 112  # Access ControlType constant for subform

_MDE_EXTENSIONS = (".mde", ".mdb", ".accdb")

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
def resolve_patient_folder(patient: dict) -> Path | None:
    """
    Cherche le dossier patient existant via son code dans le répertoire de l'instance.
    Ne crée jamais de dossier. Retourne None en cas d'échec.
    """
    inst = patient.get("instance")
    if inst is None:
        log.error("resolve_patient_folder: no instance in patient dict.")
        return None

    code = patient["code"]
    is_negative = code.startswith("-")
    digits = code[1:] if is_negative else code

    if is_negative:
        parent_dir = inst.dest_photos / f"-{digits[:1]}.000"
    else:
        parent_dir = inst.dest_photos / f"{digits[:2]}.000"

    if not parent_dir.is_dir():
        log.warning(f"[{inst.name}] Dossier parent introuvable pour le code {code}: {parent_dir}")
        return None

    for entry in parent_dir.iterdir():
        if not entry.is_dir():
            continue
        
        # On ignore un éventuel "-" parasite en tête du nom de dossier
        name_digits = entry.name.lstrip("-")
        if not name_digits.startswith(digits):
            continue
            
        # Le caractère qui suit le code ne doit pas être un chiffre
        suffix = name_digits[len(digits):]
        if suffix and suffix[0].isdigit():
            continue
            
        log.info(f"[{inst.name}] Dossier patient trouvé: {entry}")
        return entry

    log.warning(f"[{inst.name}] Aucun dossier patient trouvé pour le code {code} dans {parent_dir}")
    return None


# ROT — Access instance enumeration
def _instance_from_mde_moniker(mde_path: str) -> "Instance | None":
    """
    Exact match between a ROT moniker path and inst.mde_path (case-insensitive,
    normalised slashes). Uses == instead of 'in' to avoid 'Studiov2000' matching
    'Studiov2000-OM' as a substring.
    """
    norm = mde_path.lower().replace("/", "\\").strip("\\")
    for inst in INSTANCES:
        c = str(inst.mde_path).lower().replace("/", "\\").strip("\\")
        if c == norm:
            return inst
    return None


def _get_all_access_apps() -> list[tuple[str, object, "Instance | None"]]:
    """
    Returns all active Access.Application objects found in the ROT.
    Each entry: (mde_path, access_app, instance_or_None).
    StudioVision registers under the local .mde path, not 'Access.Application'.
    pywin32 may return enum.Next() as a tuple — unwrapped accordingly.
    """
    result = []
    try:
        rot  = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            try:
                item = enum.Next()
            except Exception:
                break
            if item is None:
                break

            moniker = item[0] if isinstance(item, tuple) else item

            try:
                ctx  = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
            except Exception:
                continue

            if not any(name.lower().endswith(ext) for ext in _MDE_EXTENSIONS):
                continue

            try:
                obj      = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                db_obj   = win32com.client.Dispatch(dispatch)
                try:
                    app = db_obj.Application
                except Exception:
                    app = db_obj
                _ = app.Forms.Count  # verify it is a valid Access application
                inst = _instance_from_mde_moniker(name)
                result.append((name, app, inst))
                log.debug(f"ROT: [{name}] -> instance={inst.name if inst else 'unknown'}")
            except Exception as exc:
                log.debug(f"ROT: could not bind [{name}]: {exc}")

    except Exception as exc:
        log.debug(f"ROT enumeration error: {exc}")

    return result


# Active patient detection
def _extract_patient_fields(form) -> dict | None:
    """
    Reads the 3 patient identity fields directly from an Access form.
    Does not recurse into subforms to avoid false matches from MENU GENERAL.
    """
    target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
    data: dict = {}
    try:
        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                name = str(ctrl.Name)
                if name in target:
                    data[name] = ctrl.Value
            except Exception:
                pass
    except Exception:
        pass
    if not target.issubset(data.keys()):
        return None
    return {
        "code":   str(data[ACCESS_FIELD_CODE]),
        "nom":    str(data[ACCESS_FIELD_NOM]),
        "prenom": str(data[ACCESS_FIELD_PRENOM]),
    }


def _read_patient_from_access(access_app) -> dict | None:
    """Returns the patient open in fPATIENTS, or None if the form is not open."""
    try:
        fc = access_app.Forms.Count
        for i in range(fc):
            try:
                form = access_app.Forms(i)
                if form.Name != PATIENT_FORM_NAME:
                    continue
                return _extract_patient_fields(form)
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_active_patient() -> dict | None:
    """
    Returns the patient with fPATIENTS open across all Access instances.
    - Scans all instances via the ROT.
    - If one instance has fPATIENTS open, returns that patient.
    - If multiple, picks the foreground instance.
    - Falls back to GetActiveObject for single-instance setups.
    """
    if not WIN32_AVAILABLE:
        return None

    fg_pid = None
    try:
        fg_hwnd = win32gui.GetForegroundWindow()
        _, fg_pid = win32process.GetWindowThreadProcessId(fg_hwnd)
    except Exception:
        pass

    candidates = []  # (is_fg, patient, inst)

    for mde_path, access_app, inst in _get_all_access_apps():
        patient = _read_patient_from_access(access_app)
        if patient is None:
            log.debug(f"ROT: no patient form open in [{mde_path}]")
            continue

        patient["instance"] = inst

        is_fg = False
        try:
            app_hwnd = access_app.hWndAccessApp()
            _, app_pid = win32process.GetWindowThreadProcessId(app_hwnd)
            is_fg = (fg_pid is not None and app_pid == fg_pid)
        except Exception:
            pass

        log.debug(
            f"ROT: patient found — code={patient['code']}  "
            f"instance={inst.name if inst else 'unknown'}  fg={is_fg}  [{mde_path}]"
        )

        if inst is None:
            log.warning(f"ROT: unknown instance for [{mde_path}] — check mde_path in INSTANCES")

        candidates.append((is_fg, patient, inst))

    if len(candidates) == 1:
        _, patient, inst = candidates[0]
        log.info(
            f"Patient selected: {patient['code']}  "
            f"instance={inst.name if inst else 'unknown'}  (only one open)"
        )
        return patient

    if len(candidates) > 1:
        candidates.sort(key=lambda x: x[0], reverse=True)
        is_fg, patient, inst = candidates[0]
        log.info(
            f"Patient selected: {patient['code']}  "
            f"instance={inst.name if inst else 'unknown'}  "
            f"({'foreground' if is_fg else 'first found — ambiguous'})"
        )
        return patient

    # Fallback: GetActiveObject (single-instance)
    log.debug("ROT: no patient found — falling back to GetActiveObject.")
    try:
        access  = win32com.client.GetActiveObject("Access.Application")
        patient = _read_patient_from_access(access)
        if patient is None:
            return None
        patient["instance"] = None
        log.debug(f"Fallback COM: patient {patient['code']}")
        return patient
    except Exception as exc:
        log.debug(f"COM GetActiveObject error: {exc}")
        return None


# Access COM — SFDoc lookup and insertion
def _find_sfdoc(form) -> object | None:
    """Recursively searches the form control tree for the SFDoc subform."""
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


def _find_patient_form_in_app(access_app) -> object | None:
    """Returns the fPATIENTS form if open in the given Access instance, otherwise None."""
    try:
        for i in range(access_app.Forms.Count):
            try:
                form = access_app.Forms(i)
                if form.Name == PATIENT_FORM_NAME:
                    return form
            except Exception:
                continue
    except Exception:
        pass
    return None


def gui_insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """
    Inserts a new record into SFDoc in the Access instance identified by
    patient["instance"]. Locates the instance via the ROT to target the
    correct window regardless of focus.
    Returns True on success, False on any COM error.
    """
    if not WIN32_AVAILABLE:
        log.error("win32com not available — GUI insertion skipped.")
        return False

    inst = patient.get("instance")

    # Find the Access object for this instance in the ROT
    access_app = None
    if inst is not None:
        for mde_path, app, app_inst in _get_all_access_apps():
            if app_inst is not None and app_inst.name == inst.name:
                access_app = app
                break

    # Single-instance fallback if the instance is not identified
    if access_app is None:
        log.warning("GUI insert: instance not found in ROT — falling back to GetActiveObject.")
        try:
            access_app = win32com.client.GetActiveObject("Access.Application")
        except Exception as exc:
            log.error(f"GUI insert: GetActiveObject failed: {exc}")
            return False

    # Verify the same patient is still open
    current_patient = _read_patient_from_access(access_app)
    if not current_patient or current_patient["code"] != patient["code"]:
        log.warning(
            f"GUI insert aborted: patient changed "
            f"(expected={patient['code']}, "
            f"current={current_patient['code'] if current_patient else 'none'})."
        )
        return False

    # Find fPATIENTS then SFDoc in this instance
    form = _find_patient_form_in_app(access_app)
    if form is None:
        log.error(f"GUI insert: form '{PATIENT_FORM_NAME}' not found.")
        return False

    sfdoc = _find_sfdoc(form)
    if sfdoc is None:
        log.error(f"GUI insert: subform '{SFDOC_SUBFORM_NAME}' not found.")
        return False

    try:
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


def move_file(source: Path, dest_folder: Path, label: str = "") -> Path | None:
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


def prevent_sleep() -> None:
    """Prevents Windows from sleeping while the router is active."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
        log.info("Sleep prevention active.")
    except Exception as exc:
        log.warning(f"Could not set execution state: {exc}")


# Startup scan and periodic catchup
def _scan_source_for_missed_files(file_queue: queue.Queue) -> None:
    """Enqueues files present in SOURCE_DIR that were missed during downtime."""
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
    For each file: wait for unlock → wait for open patient → move file → GUI insert.
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
                    orphan_file(file)
                    _notify("Orphan file", file.name)
                    log_doctor(
                        f"ATTENTION: '{file.name}' mis en orphelin "
                        f"(aucun patient ouvert après {PATIENT_WAIT_TIMEOUT // 60} min)."
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

            inst = patient.get("instance")
            log.info(
                f"Patient: {patient['nom']} {patient['prenom']} "
                f"(code {patient['code']})  "
                f"instance={inst.name if inst else 'unknown'}"
            )

            patient_folder = resolve_patient_folder(patient)
            if not patient_folder:
                log.error(
                    f"Could not resolve folder for patient {patient['code']}. Orphaning."
                )
                orphan_file(file)
                _notify("Orphan file", file.name)
                log_doctor(
                    f"ERREUR: Dossier introuvable pour '{file.name}' "
                    f"(patient {patient['nom']} {patient['prenom']}, "
                    f"code {patient['code']})."
                )
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            relative_path = f"\\{dest.relative_to(inst.dest_photos)}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            time.sleep(_GUI_PRE_INSERT_DELAY)

            if gui_insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.info(f"Record inserted: '{dest.name}' -> {relative_path}")
                log_doctor(
                    f"SUCCÈS: '{dest.name}' ajouté au dossier de "
                    f"{patient['nom']} {patient['prenom']} (code {patient['code']})."
                )
            else:
                log.error(
                    f"GUI insertion failed for '{dest.name}' "
                    f"(patient {patient['code']}). File moved to: {dest}. Manual entry required."
                )
                _notify("Insertion error", f"'{dest.name}' moved but not inserted — see log.")
                log_doctor(
                    f"ERREUR: '{dest.name}' déplacé vers {dest} mais NON inséré "
                    f"pour le patient {patient['nom']} {patient['prenom']} "
                    f"(code {patient['code']}). Saisie manuelle requise."
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
                log.warning("Observer stopped. Attempting reconnect...")
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


# Studio Vision lifecycle
_SV_POLL_INTERVAL   = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT = 30  # Seconds to wait for msaccess.exe to appear after launch


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


_PIPE_NAME = r"\\.\pipe\ImageRouter_StudioVision_Multi_v6"


def _launch_instance(inst: Instance) -> None:
    """Launches the StudioVision instance if not already running."""
    # Check if this instance is already running via the ROT
    for mde_path, _, rot_inst in _get_all_access_apps():
        if rot_inst is not None and rot_inst.name == inst.name:
            log.info(f"[{inst.name}] Already running — skipping launch.")
            return

    log.info(f"[{inst.name}] Launching: {inst.studio_vision_cmd[0]}")
    pids_before = _get_msaccess_pids()
    try:
        subprocess.Popen(inst.studio_vision_cmd)
    except FileNotFoundError:
        log.error(f"[{inst.name}] Executable not found: {inst.studio_vision_cmd[0]}")
        return
    except Exception as exc:
        log.error(f"[{inst.name}] Launch failed: {exc}")
        return

    # Wait for msaccess.exe to appear
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT
    while time.monotonic() < deadline:
        if _get_msaccess_pids() - pids_before:
            log.info(f"[{inst.name}] StudioVision is starting.")
            return
        time.sleep(1)
    log.warning(f"[{inst.name}] msaccess.exe did not appear within timeout.")


def _pipe_server() -> None:
    """
    Named pipe server — receives instance aliases sent by secondary invocations.
    Each message is an alias string (e.g. 'HR' or 'OM').
    """
    log.info("Pipe server started.")
    while not _stop_event.is_set():
        try:
            pipe = win32pipe.CreateNamedPipe(
                _PIPE_NAME,
                win32pipe.PIPE_ACCESS_INBOUND,
                win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                win32pipe.PIPE_UNLIMITED_INSTANCES,
                256, 256,
                0, None,
            )
            # Wait for a client connection (1s timeout to remain responsive to _stop_event)
            try:
                win32pipe.ConnectNamedPipe(pipe, None)
            except pywintypes.error:
                win32file.CloseHandle(pipe)
                continue

            try:
                _, data = win32file.ReadFile(pipe, 256)
                alias = data.decode("utf-8").strip().upper()
                log.info(f"Pipe: alias received = {alias!r}")
                inst = _INSTANCE_BY_ALIAS.get(alias)
                if inst:
                    threading.Thread(
                        target=_launch_instance, args=(inst,),
                        name=f"Launch-{inst.name}", daemon=True
                    ).start()
                else:
                    log.warning(f"Pipe: unknown alias {alias!r}")
            except Exception as exc:
                log.debug(f"Pipe: read error: {exc}")
            finally:
                win32file.CloseHandle(pipe)

        except Exception as exc:
            if not _stop_event.is_set():
                log.debug(f"Pipe server error: {exc}")
            time.sleep(1)

    log.info("Pipe server stopped.")


def _pipe_send_alias(alias: str) -> bool:
    """Sends an alias to the already-running instance via named pipe. Returns True on success."""
    try:
        handle = win32file.CreateFile(
            _PIPE_NAME,
            win32file.GENERIC_WRITE,
            0, None,
            win32file.OPEN_EXISTING,
            0, None,
        )
        win32file.WriteFile(handle, alias.encode("utf-8"))
        win32file.CloseHandle(handle)
        log.info(f"Alias {alias!r} sent to running instance.")
        return True
    except Exception as exc:
        log.debug(f"Pipe send failed: {exc}")
        return False


def _launch_studio_vision(requested_alias: str | None) -> None:
    """
    Launches the requested instance and monitors all msaccess.exe processes.
    Triggers shutdown when all instances are closed.
    Force-kills zombie processes on exit to release COM locks.
    """
    pids_before: set[int] = _get_msaccess_pids()

    # Launch the requested instance at startup
    if requested_alias:
        inst = _INSTANCE_BY_ALIAS.get(requested_alias.upper())
        if inst:
            _launch_instance(inst)
        else:
            log.warning(f"Unknown alias at startup: {requested_alias!r}")
            log.warning(f"Available aliases: {list(_INSTANCE_BY_ALIAS.keys())}")

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
                log.debug(f"No StudioVision active ({consecutive_empty}/{_EMPTY_THRESHOLD}).")
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    log.info("All StudioVision instances closed. Initiating shutdown.")
                    break
            else:
                consecutive_empty = 0

    except Exception as exc:
        log.error(f"Error monitoring msaccess.exe: {exc}")

    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    log.info(f"Force-killed zombie msaccess.exe (PID {pid}).")
            except Exception:
                pass

        _KILL_DRAIN_TIMEOUT = 10
        _KILL_DRAIN_POLL    = 0.5
        deadline = time.monotonic() + _KILL_DRAIN_TIMEOUT
        while time.monotonic() < deadline:
            still_alive = {pid for pid in tracked_pids if psutil.pid_exists(pid)}
            if not still_alive:
                break
            time.sleep(_KILL_DRAIN_POLL)

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass


# Entry point
def main() -> None:
    global _icon, _mutex_handle

    # Optional argument: instance alias to launch (e.g. "HR" or "OM")
    requested_alias = sys.argv[1].upper() if len(sys.argv) > 1 else None

    log.info(f"Starting Box CV v6  argv={sys.argv}  alias={requested_alias}")

    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Multi_v6_Mutex")
    already_running = (win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS)

    if already_running:
        # Already running: forward the alias via pipe
        if requested_alias:
            if _pipe_send_alias(requested_alias):
                log.info(f"Already running — alias {requested_alias!r} forwarded.")
            else:
                log.warning(
                    f"Already running but pipe did not respond. "
                    f"{requested_alias} may already be open."
                )
        else:
            log.info("Already running and no alias provided — nothing to do.")
        sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Studio Vision Image Router — Version 6 Multi-Instance ===")
    log.info(f"  Source     : {SOURCE_DIR}")
    log.info(f"  Orphans    : {ORPHAN_DIR}")
    for inst in INSTANCES:
        log.info(f"  [{inst.name}] alias={inst.alias}  mde={inst.mde_path}  photos={inst.dest_photos}")
    log.info(f"  Aliases    : {list(_INSTANCE_BY_ALIAS.keys())}")
    log.info(f"  Requested  : {requested_alias or '(none)'}")
    log.info(f"  Timeout    : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Catchup    : every {CATCHUP_INTERVAL}s")
    log.info(f"  Doctor log : {_get_doctor_log_path()}")

    log_doctor(f"Routeur d'images démarré (v6 multi). Surveille: {SOURCE_DIR}")

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(target=worker,          args=(file_queue,), name="Worker",     daemon=True).start()
    threading.Thread(target=_catchup_loop,   args=(file_queue,), name="Catchup",    daemon=True).start()
    threading.Thread(target=_run_background, args=(file_queue,), name="Background", daemon=True).start()
    threading.Thread(target=_pipe_server,    name="PipeServer",                     daemon=True).start()

    log.info("Startup scan — checking for pending files...")
    _scan_source_for_missed_files(file_queue)

    threading.Thread(
        target=_launch_studio_vision, args=(requested_alias,),
        name="StudioVisionLauncher", daemon=True
    ).start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running headless.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested via keyboard.")
        finally:
            _stop_event.set()
        log.info("Application stopped.")
        log_doctor("Routeur d'images arrêté.")
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
    log_doctor("Routeur d'images arrêté.")


if __name__ == "__main__":
    main()