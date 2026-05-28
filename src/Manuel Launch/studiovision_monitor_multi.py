"""
Routes image files to the correct patient folder, inserts a record
into the Access database, and refreshes the interface.

VERSION PIERRE-HENRI (Box-6)

• Two desktop shortcuts each launch THIS program + the associated StudioVision:
      lancer_OM.bat   →  StudioVision OM  + database \\studiovision\Studiov2000-OM
      lancer_HR.bat   →  StudioVision HR  + database \\studiovision\Studiov2000

• Only one Python process runs at a time (Windows mutex).

• NORMAL MODE: only one shortcut clicked → automatic send to the chosen DB.

• MANUAL MODE: both shortcuts clicked (both SV are open)
  → notification bottom right + systray menu to choose the target DB.
  → "↩ Quit manual mode" button to return to normal mode.

• SOURCE_DIR and ORPHAN_DIR are SHARED between both instances.

Usage (from the provided .bat files):
    pythonw studiovision_monitor_multi.py --mode OM
    pythonw studiovision_monitor_multi.py --mode HR

Dependencies: watchdog, pyodbc, pywin32, pythoncom, pystray, Pillow, psutil
"""

import argparse
import ctypes
import logging
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import pythoncom
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

import win32api
import win32event
import winerror
import psutil


BOX_NAME = "StudioVision Box-6"

SOURCE_DIR = Path(r"C:\Users\Box-6\Desktop\Export CV")
ORPHAN_DIR = Path(r"C:\Users\Box-6\Desktop\Images_Oubliées")

WATCHED_EXTENSIONS     = {".jpg", ".jpeg", ".jfif", ".png", ".bmp",
                           ".tif", ".tiff", ".dcm", ".pdf", ".rtf",
                           ".doc", ".docx", ".odt"}
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3
PATIENT_WAIT_TIMEOUT   = 900   # seconds (15 min)

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"
SFDOC_SUBFORM_NAME  = "SFDoc"


@dataclass
class Instance:
    name:             str
    exe_path:         str
    launch_args:      list
    dest_photos:      Path
    public_mdb:       Path
    docum_mdb:        Path
    exam_description: dict


INSTANCES: list[Instance] = [
    Instance(
        name        = "OM",
        exe_path    = r"C:\Studiov2000-OM\svprog\msaccess.exe",
        launch_args = [
            "/runtime", r"C:\Studiov2000-OM\svprog\Ophprog.mde",
            "/wrkgrp",  r"C:\Studiov2000-OM\config\system.mdw",
            "/User", "/Pwd", "/X", "demarrage",
        ],
        dest_photos = Path(r"\\studiovision\Studiov2000-OM\PHOTOS"),
        public_mdb  = Path(r"\\studiovision\Studiov2000-OM\fichier\PUBLIC.MDB"),
        docum_mdb   = Path(r"\\studiovision\Studiov2000-OM\fichier\DOCUM.MDB"),
        exam_description = {
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
        },
    ),
    Instance(
        name        = "HR",
        exe_path    = r"C:\Studiov2000\Svprog\MSACCESS.EXE",
        launch_args = [
            "/runtime", r"C:\Studiov2000\svprog\Ophprog.mde",
            "/wrkgrp",  r"C:\Studiov2000\config\system.mdw",
            "/User", "/Pwd", "/X", "demarrage",
        ],
        dest_photos = Path(r"\\studiovision\Studiov2000\PHOTOS"),
        public_mdb  = Path(r"\\studiovision\Studiov2000\fichier\PUBLIC.MDB"),
        docum_mdb   = Path(r"\\studiovision\Studiov2000\fichier\DOCUM.MDB"),
        exam_description = {
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
        },
    ),
]

INSTANCES_BY_NAME: dict[str, Instance] = {inst.name: inst for inst in INSTANCES}

_SIGNAL_DIR  = Path(tempfile.gettempdir())
_SIGNAL_FILE = {
    inst.name: _SIGNAL_DIR / f"sv_signal_{inst.name}.tmp"
    for inst in INSTANCES
}

_MUTEX_NAME = "ImageRouter_StudioVision_PierreHenri_Box6_Mutex"

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


_NETWORK_SHARE_POLL = 10

_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)
_COLOR_MANUAL = (255, 140, 0)
_ICON_SIZE    = 64

_icon:         "pystray.Icon | None" = None
_status_text:  str                   = "Démarrage..."
_stop_event:   threading.Event       = threading.Event()
_mutex_handle                        = None

_state_lock:             threading.Lock = threading.Lock()
_selected_instance_name: str           = INSTANCES[0].name
_manual_mode:            bool          = False


def get_selected_instance() -> Instance:
    """Thread-safe getter for the currently selected instance."""
    with _state_lock:
        return INSTANCES_BY_NAME[_selected_instance_name]


def _is_manual_mode() -> bool:
    with _state_lock:
        return _manual_mode


def set_selected_instance(name: str) -> None:
    """Change the target DB and update status (thread-safe)."""
    global _selected_instance_name
    with _state_lock:
        _selected_instance_name = name
        manual = _manual_mode
    log.info(f"Selected DB: [{name}]")
    prefix = "⚡ MODE MANUEL" if manual else BOX_NAME
    _set_status(f"{prefix} — BDD : {name}")
    _refresh_menu()


def activate_manual_mode() -> None:
    """
    Switch to manual mode: both StudioVision are open,
    user chooses which DB to send to.
    """
    global _manual_mode
    with _state_lock:
        if _manual_mode:
            return   # already active
        _manual_mode = True
        inst_name = _selected_instance_name

    log.info("Manual mode activated (both StudioVision detected).")
    _set_status(f"⚡ MODE MANUEL — BDD : {inst_name}")
    _notify(
        "⚡ Mode Manuel Activé",
        "Both StudioVision are open.\nClick the blue icon bottom right to choose the database.",
    )
    _refresh_menu()


def deactivate_manual_mode(icon=None, item=None) -> None:
    """Return to normal mode (only one SV / one DB)."""
    global _manual_mode
    with _state_lock:
        _manual_mode = False
        inst_name = _selected_instance_name

    log.info(f"Manual mode deactivated. Active DB: {inst_name}")
    _set_status(f"{BOX_NAME} — BDD : {inst_name} — Prêt")
    _notify("Mode Normal", f"Active database: {inst_name}")
    _refresh_menu()


def _make_icon_image(color: tuple) -> "Image.Image":
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    m    = 4
    draw.ellipse([m, m, _ICON_SIZE - m, _ICON_SIZE - m], fill=color)
    return img


def _set_status(text: str, processing: bool = False) -> None:
    """Update status text and systray icon color."""
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            manual = _is_manual_mode()
            if manual:
                color = _COLOR_MANUAL
            elif processing:
                color = _COLOR_ACTIVE
            else:
                color = _COLOR_READY
            _icon.icon = _make_icon_image(color)
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Icon update failed: {e}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            log.debug(f"Notification failed: {e}")


def _open_logs(icon=None, item=None) -> None:
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning(f"Could not open logs: {e}")


def _quit(icon=None, item=None) -> None:
    log.info("Exit requested from systray menu.")
    _stop_event.set()
    if _icon is not None:
        _icon.stop()


class _InstanceMenuHandler:
    """Handler per StudioVision instance for systray menu."""
    def __init__(self, name: str) -> None:
        self._name = name

    def action(self, icon=None, item=None) -> None:
        set_selected_instance(self._name)

    def checked(self, item=None) -> bool:
        with _state_lock:
            return _selected_instance_name == self._name


_INSTANCE_HANDLERS: dict[str, _InstanceMenuHandler] = {
    inst.name: _InstanceMenuHandler(inst.name) for inst in INSTANCES
}


def _build_menu() -> "pystray.Menu | None":
    """
    Build systray menu according to current state:
    - Normal mode: submenu "Database" with radio selection
    - Manual mode: direct OM / HR buttons + "Quit manual mode" button
    """
    if not TRAY_AVAILABLE:
        return None

    status_item = pystray.MenuItem(
        lambda item: _status_text, action=None, enabled=False,
    )
    sep   = pystray.Menu.SEPARATOR
    logs  = pystray.MenuItem("📋 Ouvrir les logs", _open_logs)
    quit_ = pystray.MenuItem("✖ Quitter le programme", _quit)

    with _state_lock:
        manual = _manual_mode

    if manual:
        instance_items = [
            pystray.MenuItem(
                f"📤  Envoyer vers {inst.name}",
                action=_INSTANCE_HANDLERS[inst.name].action,
                checked=_INSTANCE_HANDLERS[inst.name].checked,
                radio=True,
            )
            for inst in INSTANCES
        ]
        return pystray.Menu(
            status_item,
            sep,
            *instance_items,
            sep,
            pystray.MenuItem("↩  Quitter mode manuel", deactivate_manual_mode),
            sep,
            logs,
            quit_,
        )
    else:
        submenu_items = [
            pystray.MenuItem(
                f"Base : {inst.name}",
                action=_INSTANCE_HANDLERS[inst.name].action,
                checked=_INSTANCE_HANDLERS[inst.name].checked,
                radio=True,
            )
            for inst in INSTANCES
        ]
        return pystray.Menu(
            status_item,
            sep,
            pystray.MenuItem("🗄  Base de données", pystray.Menu(*submenu_items)),
            sep,
            logs,
            quit_,
        )


def _refresh_menu() -> None:
    """Rebuild and apply systray menu (thread-safe)."""
    if _icon is not None:
        try:
            _icon.menu = _build_menu()
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Menu rebuild failed: {e}")


def _check_signals() -> None:
    """
    Called every second by the background thread.
    Detects if a 2nd shortcut was clicked (signal file in %TEMP%)
    and activates manual mode accordingly.
    """
    for inst_name, signal_file in _SIGNAL_FILE.items():
        if signal_file.exists():
            deleted = False
            try:
                signal_file.unlink()
                deleted = True
            except Exception as e:
                log.warning(f"Could not delete signal file {signal_file}: {e}")

            if deleted:
                log.info(f"Signal received: shortcut {inst_name} clicked while program already running.")
                activate_manual_mode()


def _launch_studiovision(inst: Instance) -> None:
    """Launch the StudioVision associated with the given instance."""
    cmd = [inst.exe_path] + inst.launch_args
    try:
        subprocess.Popen(cmd, close_fds=True)
        log.info(f"[{inst.name}] StudioVision launched: {inst.exe_path}")
    except FileNotFoundError:
        log.error(f"[{inst.name}] Executable not found: {inst.exe_path}")
        ctypes.windll.user32.MessageBoxW(
            0,
            f"StudioVision {inst.name} not found:\n\n"
            f"{inst.exe_path}\n\n"
            "Check the path in the script (exe_path variable).",
            f"Error — StudioVision {inst.name}",
            0x10,
        )
    except Exception as e:
        log.error(f"[{inst.name}] Could not launch StudioVision: {e}")


def wait_for_network_share() -> None:
    is_network = str(SOURCE_DIR).startswith("\\\\") or str(SOURCE_DIR).startswith("//")
    if not is_network:
        return
    attempt = 0
    while not SOURCE_DIR.is_dir():
        attempt += 1
        log.warning(
            f"Network share inaccessible: {SOURCE_DIR}  "
            f"(attempt {attempt}, retry in {_NETWORK_SHARE_POLL}s)"
        )
        time.sleep(_NETWORK_SHARE_POLL)
    if attempt:
        log.info(f"Network share accessible after {attempt} attempt(s): {SOURCE_DIR}")


def db_connect(mdb_path: Path):
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def find_patient_folder(patient_code: str, instance: Instance) -> "Path | None":
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available.")
        return None
    if not instance.public_mdb.exists():
        log.error(f"[{instance.name}] PUBLIC.MDB not found: {instance.public_mdb}")
        return None
    try:
        conn   = db_connect(instance.public_mdb)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            log.warning(f"[{instance.name}] No existing document for patient {patient_code}.")
            return None

        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error(f"[{instance.name}] Unexpected 'Photo externe' format: {row[0]}")
            return None

        folder = instance.dest_photos / parts[0] / parts[1]
        if not folder.is_dir():
            log.error(f"[{instance.name}] Patient folder not found: {folder}")
            return None

        log.info(f"[{instance.name}] Patient folder resolved: {folder}")
        return folder
    except Exception as e:
        log.error(f"[{instance.name}] DB error (find_patient_folder): {e}")
        return None


def insert_document(
    patient: dict, relative_path: str, description: str, instance: Instance
) -> bool:
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert ignored.")
        return False
    if not instance.public_mdb.exists():
        log.error(f"[{instance.name}] PUBLIC.MDB not found, insert ignored.")
        return False
    try:
        conn   = db_connect(instance.public_mdb)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO Documents
                ([code patient], [Date], DESCRIPTIONS, TEXTE,
                 [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, ?, 99, NULL)
            """,
            (int(patient["code"]), datetime.now(), description,
             relative_path, relative_path)
        )
        conn.commit()
        conn.close()
        log.info(
            f"[{instance.name}] Insert OK: patient={patient['code']} "
            f"path='{relative_path}'"
        )
        return True
    except Exception as e:
        log.error(f"[{instance.name}] DB insert failed: {e}")
        return False


_AC_SUBFORM = 112


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


def _read_patient_from_access(access_app) -> "dict | None":
    try:
        form = access_app.Screen.ActiveForm
    except Exception:
        return None
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


def get_active_patient() -> "dict | None":
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        log.debug(f"GetActiveObject failed: {e}")
        return None
    patient = _read_patient_from_access(access)
    if patient:
        log.info(
            f"Active patient: {patient['nom']} {patient['prenom']} "
            f"(code {patient['code']})"
        )
    return patient


def refresh_ui(expected_patient_code: "str | None" = None) -> None:
    """Refresh the active Access form (SFDoc Requery + MoveLast)."""
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh ignored: no active form.")
            return

        if expected_patient_code is not None:
            current      = get_active_patient()
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    f"Refresh ignored: patient changed "
                    f"(expected={expected_patient_code}, current={current_code})."
                )
                return

        try:
            form.Refresh()
            log.info(f"Refresh() on form '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() failed ({e_ref}), continuing...")

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found. "
                "SFDoc refresh ignored."
            )
            return

        try:
            if form.Dirty:
                log.info("Form in edit mode (Dirty=True); clearing before Requery.")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Dirty check/clear failed: {e_dirty}")

        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5
        requery_ok        = False

        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}' (attempt {attempt})")
                requery_ok = True
                break
            except Exception as e_req:
                log.warning(
                    f"Requery() attempt {attempt}/{_REQUERY_ATTEMPTS} failed: {e_req}"
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requery_ok:
            try:
                sfdoc.Refresh()
                log.info(f"Fallback Refresh() on '{SFDOC_SUBFORM_NAME}'")
            except Exception as e_ref2:
                log.warning(f"Fallback Refresh() failed: {e_ref2}")

        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed: {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


def wait_for_file(file: Path) -> bool:
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"File locked ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retry...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error(
        f"File still locked after {FILE_LOCK_MAX_ATTEMPTS} attempts: {file}"
    )
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> "Path | None":
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict, renamed to {dest.name}")
    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} → {dest}")
        return dest
    except Exception as e:
        log.error(f"Move failed: {e}")
        return None


def orphan_file(file: Path) -> None:
    log.warning(f"Orphan file: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def prevent_sleep() -> None:
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention enabled.")
    except Exception as e:
        log.warning(f"SetThreadExecutionState failed: {e}")


def worker(file_queue: queue.Queue) -> None:
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh:     bool      = False
    last_patient_code: str|None  = None
    burst_count:       int       = 0

    try:
        while True:
            # Burst debounce wait 1.5s
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                # End of burst: flush UI refresh
                if needs_refresh:
                    instance = get_selected_instance()
                    msg = (
                        f"Send to DB {instance.name} "
                        f"({burst_count} file(s))"
                    )
                    log.info(f"Burst complete — {msg}")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh     = False
                    last_patient_code = None
                    burst_count       = 0
                    _notify(f"DB {instance.name} — Transfer complete", msg)
                    prefix = "⚡ MODE MANUEL" if _is_manual_mode() else BOX_NAME
                    _set_status(f"{prefix} — BDD : {instance.name} — Prêt")
                continue
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue

            instance = get_selected_instance()
            log.info(
                f"Processing: {file.name} ({file_queue.qsize()} pending) "
                f"→ DB [{instance.name}]"
            )

            if burst_count == 0 and not needs_refresh:
                _notify(f"DB {instance.name} — Transfer in progress", file.name)
            _set_status(
                f"Sending to DB {instance.name}...", processing=True
            )

            if not file.exists():
                log.warning(f"File disappeared before processing: {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Abandon — persistent lock: {file.name}")
                _notify("Error", f"File locked: {file.name}")
                file_queue.task_done()
                continue

            # Wait for active patient
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
                f"(code {patient['code']}) → DB [{instance.name}]"
            )

            patient_folder = find_patient_folder(patient["code"], instance)
            if not patient_folder:
                log.error(
                    f"[{instance.name}] Patient folder not found for "
                    f"{patient['code']}. Orphaned."
                )
                orphan_file(file)
                _notify("Orphan file", file.name)
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = instance.exam_description.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description, instance):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.debug(
                    f"Insert OK in [{instance.name}] — deferred refresh."
                )
            else:
                log.warning("Insert failed.")
                _notify("DB Error", "Insert failed — check logs")

            file_queue.task_done()

    finally:
        if needs_refresh:
            instance = get_selected_instance()
            log.info(
                f"Worker stopping — flush refresh (DB [{instance.name}])."
            )
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                _notify(
                    f"DB {instance.name} — Transfer complete",
                    f"Send to DB {instance.name} ({burst_count} file(s))",
                )
        _set_status(f"{BOX_NAME} — Stopped")
        pythoncom.CoUninitialize()


class ImageProducer(FileSystemEventHandler):
    def __init__(self, fq: queue.Queue) -> None:
        super().__init__()
        self._queue = fq

    def on_created(self, event) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info(f"Enqueued : {file.name} (queue : {self._queue.qsize() + 1})")
        self._queue.put(file)


def _run_background(file_queue: queue.Queue) -> None:
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — monitoring source folder.")
        return obs

    observer = _start_observer()
    inst     = get_selected_instance()
    _set_status(f"{BOX_NAME} — BDD : {inst.name} — Prêt")

    try:
        while not _stop_event.is_set():
            _check_signals()
            if not observer.is_alive():
                log.warning("Observer stopped (network loss?). Reconnecting...")
                _set_status(f"{BOX_NAME} — Reconnecting...")
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                log.info(
                    f"Waiting {_RECONNECT_WAIT}s before restarting observer..."
                )
                time.sleep(_RECONNECT_WAIT)
                observer = _start_observer()
                inst     = get_selected_instance()
                if not _is_manual_mode():
                    _set_status(f"{BOX_NAME} — BDD : {inst.name} — Prêt")
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


def main() -> None:
    global _icon, _mutex_handle, _selected_instance_name

    parser = argparse.ArgumentParser(
        description=f"{BOX_NAME} — Image router"
    )
    parser.add_argument(
        "--mode",
        choices=[inst.name for inst in INSTANCES],
        default=INSTANCES[0].name,
        help="StudioVision instance to target at startup (OM or HR)",
    )
    args = parser.parse_args()
    mode = args.mode

    _mutex_handle = win32event.CreateMutex(None, False, _MUTEX_NAME)
    already_running = (win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS)

    if already_running:
        log.info(
            f"Program already running. "
            f"Sending signal [{mode}] and launching StudioVision {mode}."
        )
        try:
            _SIGNAL_FILE[mode].touch()
            log.info(f"Signal file created: {_SIGNAL_FILE[mode]}")
        except Exception as e:
            log.error(f"Could not create signal file: {e}")

        _launch_studiovision(INSTANCES_BY_NAME[mode])
        sys.exit(0)

    _selected_instance_name = mode
    inst                    = INSTANCES_BY_NAME[mode]

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        ctypes.windll.user32.MessageBoxW(
            0,
            f"Source folder not found:\n{SOURCE_DIR}",
            "Error — Image router",
            0x10,
        )
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    for sf in _SIGNAL_FILE.values():
        try:
            sf.unlink(missing_ok=True)
        except Exception:
            pass

    log.info("═" * 60)
    log.info(f"Startup {BOX_NAME} — initial mode: [{mode}]")
    log.info(f"  Source    : {SOURCE_DIR}")
    log.info(f"  Orphans   : {ORPHAN_DIR}")
    for i in INSTANCES:
        log.info(f"  [{i.name}]  exe  : {i.exe_path}")
        log.info(f"  [{i.name}]  mdb  : {i.public_mdb}")
    log.info(f"  Timeout   : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Extensions: {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info("═" * 60)

    _launch_studiovision(inst)

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    ).start()
    threading.Thread(
        target=_run_background, args=(file_queue,), name="Background", daemon=True
    ).start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without systray icon.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Exit requested.")
        finally:
            _stop_event.set()
        return

    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=_make_icon_image(_COLOR_READY),
        title=BOX_NAME,
        menu=_build_menu(),
    )

    log.info("Systray icon started.")
    _icon.run()

    _stop_event.set()
    log.info("Application stopped.")


if __name__ == "__main__":
    main()