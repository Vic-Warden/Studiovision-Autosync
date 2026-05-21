"""
Routes incoming imaging files to the correct patient folder,
inserts a DB record, and refreshes the Access UI.

MULTI-INSTANCE VERSION: a single program monitors SOURCE_DIR and routes files
to the instance (Megret or Romoli) explicitly selected by the user via the
system tray menu. No automatic detection of the active instance is performed.

Pipeline: PollingObserver → file_queue → Worker → Access DB + UI refresh
          (1.5 s burst debounce, auto-reconnect on network drop)

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


BOX_NAME   = "StudioVision Multi"

SOURCE_DIR = Path(r"C:\Users\Box-6\Desktop\Export CV")
ORPHAN_DIR = Path(r"C:\Users\Box-6\Desktop\Images_Oubliées")

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


@dataclass
class Instance:
    name:        str
    exe:         str
    dest_photos: Path
    public_mdb:  Path
    docum_mdb:   Path


INSTANCES: list[Instance] = [
    Instance(
        name        = "Romoli",
        exe         = "msaccess.exe",
        dest_photos = Path(r"\\studiovision\Studiov2000\PHOTOS"),
        public_mdb  = Path(r"\\studiovision\Studiov2000\fichier\PUBLIC.MDB"),
        docum_mdb   = Path(r"\\studiovision\Studiov2000\fichier\DOCUM.MDB"),
    ),
    Instance(
        name        = "Megret",
        exe         = "msaccess.exe",
        dest_photos = Path(r"\\studiovision\Studiov2000-OM\PHOTOS"),
        public_mdb  = Path(r"\\studiovision\Studiov2000-OM\fichier\PUBLIC.MDB"),
        docum_mdb   = Path(r"\\studiovision\Studiov2000-OM\fichier\DOCUM.MDB"),
    ),
]

# Build a quick lookup by name for the menu handlers.
INSTANCES_BY_NAME: dict[str, Instance] = {inst.name: inst for inst in INSTANCES}

# Default instance selected at startup (Megret).
_DEFAULT_INSTANCE_NAME = "Megret"

# Used for the double-click guard: block launch if any known SV process is running.
_ALL_SV_EXES = {inst.exe.lower() for inst in INSTANCES}


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

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)

_icon: "pystray.Icon | None" = None
_status_text: str             = "Starting..."
_stop_event: threading.Event  = threading.Event()
_mutex_handle                 = None

# ── Selected instance (protected by _instance_lock) ──────────────────────────
_selected_instance_name: str      = _DEFAULT_INSTANCE_NAME
_instance_lock:          threading.Lock = threading.Lock()


def get_selected_instance() -> Instance:
    """Return the Instance currently selected in the tray menu (thread-safe)."""
    with _instance_lock:
        return INSTANCES_BY_NAME[_selected_instance_name]


def set_selected_instance(name: str) -> None:
    """Switch the active instance and update tray status (thread-safe)."""
    global _selected_instance_name
    with _instance_lock:
        _selected_instance_name = name
    log.info(f"BDD sélectionnée : [{name}]")
    _set_status(f"{BOX_NAME} — BDD : {name}", processing=False)
    if _icon is not None:
        try:
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Menu update failed: {e}")


# ── Tray helpers ──────────────────────────────────────────────────────────────

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


def _open_logs(icon, item) -> None:   # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning(f"Could not open log file: {e}")


def _quit(icon, item) -> None:        # noqa: ARG001
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


# ── Instance selector menu items ──────────────────────────────────────────────
#
# pystray _assert_action rejects closures/lambdas on older versions.
# Bound methods of a class pass the check unconditionally.

class _InstanceMenuHandler:
    """One handler per StudioVision instance; bound methods used as callbacks."""
    def __init__(self, name: str) -> None:
        self._name = name

    def action(self, icon, item) -> None:   # noqa: ARG002
        set_selected_instance(self._name)

    def checked(self, item) -> bool:        # noqa: ARG002
        with _instance_lock:
            return _selected_instance_name == self._name


# Pre-build handlers so they stay alive for the lifetime of the process.
_INSTANCE_HANDLERS: dict[str, "_InstanceMenuHandler"] = {
    inst.name: _InstanceMenuHandler(inst.name) for inst in INSTANCES
}


def _make_instance_menu() -> "pystray.Menu":
    """Radio-style sub-menu — one entry per instance, bound methods only."""
    items = [
        pystray.MenuItem(
            text=f"Base : {inst.name}",
            action=_INSTANCE_HANDLERS[inst.name].action,
            checked=_INSTANCE_HANDLERS[inst.name].checked,
            radio=True,
        )
        for inst in INSTANCES
    ]
    return pystray.Menu(*items)


# ── Network / DB helpers ──────────────────────────────────────────────────────

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
        log.info(f"Network share is now accessible after {attempt} attempt(s): {SOURCE_DIR}")


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
            log.warning(f"[{instance.name}] No existing document found for patient {patient_code}.")
            return None

        # Derive the patient folder from an existing document path stored in the DB.
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error(f"[{instance.name}] Unexpected 'Photo externe' format: {row[0]}")
            return None

        folder = instance.dest_photos / parts[0] / parts[1]
        if not folder.is_dir():
            log.error(f"[{instance.name}] Resolved folder not found on disk: {folder}")
            return None

        log.info(f"[{instance.name}] Patient folder resolved: {folder}")
        return folder
    except Exception as e:
        log.error(f"[{instance.name}] DB folder lookup failed: {e}")
        return None


def insert_document(patient: dict, relative_path: str, description: str, instance: Instance) -> bool:
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert skipped.")
        return False
    if not instance.public_mdb.exists():
        log.error(f"[{instance.name}] PUBLIC.MDB not found, insert skipped.")
        return False
    try:
        conn   = db_connect(instance.public_mdb)
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
        log.info(
            f"[{instance.name}] Insert OK: patient={patient['code']} "
            f"path='{relative_path}' db={instance.public_mdb.name}"
        )
        return True
    except Exception as e:
        log.error(f"[{instance.name}] DB insert failed: {e}")
        return False


# ControlType constant for Access subform controls.
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
    """Extract patient fields from the active form of a given Access COM object."""
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
    """
    Read the patient currently open in any running Access window.
    Returns only the patient record (code, nom, prénom).
    Instance routing is handled separately via the tray menu selection.
    """
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
    """Refresh the active Access form: SFDoc Requery + MoveLast."""
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return

        # Guard against a patient change occurring between file processing and refresh.
        if expected_patient_code is not None:
            current      = get_active_patient()
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

        # Clear Dirty state to avoid Access blocking the Requery call.
        try:
            if form.Dirty:
                log.info("Parent form is in edit mode (Dirty=True); clearing Dirty before Requery.")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Dirty check/clear failed ({e_dirty}), continuing...")

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
                log.warning(f"Fallback Refresh() also failed on '{SFDOC_SUBFORM_NAME}': {e_ref2}")

        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")

    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")


# ── File helpers ──────────────────────────────────────────────────────────────

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


def move_file(source: Path, dest_folder: Path, label: str = "") -> "Path | None":
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        # Append a timestamp to avoid silently overwriting an existing file.
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


# ── Worker ────────────────────────────────────────────────────────────────────

def prevent_sleep() -> None:
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning(f"Could not set execution state: {e}")


def worker(file_queue: queue.Queue) -> None:
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh: bool           = False
    last_patient_code: str | None = None
    burst_count: int              = 0

    try:
        while True:
            # Block for up to 1.5 s to collect burst; flush UI refresh when idle.
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    instance = get_selected_instance()
                    msg = (
                        f"Envoi dans la BDD de {instance.name} "
                        f"({burst_count} fichier(s))"
                    )
                    log.info(f"Burst complet — {msg}")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh = False
                    last_patient_code = None
                    _notify(f"BDD {instance.name} — Transfert terminé", msg)
                    _set_status(f"{BOX_NAME} — BDD : {instance.name} — Prêt", processing=False)
                    burst_count = 0
                continue
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue

            instance = get_selected_instance()
            log.info(
                f"Processing: {file.name} ({file_queue.qsize()} pending) "
                f"→ BDD [{instance.name}]"
            )

            if burst_count == 0 and not needs_refresh:
                _notify(
                    f"BDD {instance.name} — Transfert en cours",
                    file.name,
                )
            _set_status(
                f"Envoi dans la BDD de {instance.name}...",
                processing=True,
            )

            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                _notify("Erreur", f"Fichier verrouillé : {file.name}")
                file_queue.task_done()
                continue

            # ── Read the patient code from the active Access window ────────
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
                    _notify("Fichier orphelin", file.name)
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(
                        f"Aucun patient ouvert, en attente "
                        f"(timeout dans {PATIENT_WAIT_TIMEOUT // 60} min)"
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                f"Patient : {patient['nom']} {patient['prenom']} "
                f"(code {patient['code']}) → BDD [{instance.name}]"
            )

            # ── Route to the explicitly selected instance ──────────────────
            # (no automatic detection — the user's tray choice is the only source of truth)

            patient_folder = find_patient_folder(patient["code"], instance)
            if not patient_folder:
                log.error(
                    f"[{instance.name}] Could not resolve folder for patient {patient['code']}. "
                    "Orphaning."
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
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description, instance):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.debug(
                    f"Insert OK dans [{instance.name}] — "
                    "needs_refresh=True (refresh différé à la fin du burst)."
                )
            else:
                log.warning("Insert échoué, refresh flag inchangé.")
                _notify("Erreur BDD", "Insertion échouée — consultez les logs")

            file_queue.task_done()

    finally:
        if needs_refresh:
            instance = get_selected_instance()
            log.info(
                f"Worker s'arrête — flush du refresh UI en attente "
                f"(BDD [{instance.name}])."
            )
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                msg = f"Envoi dans la BDD de {instance.name} ({burst_count} fichier(s))"
                _notify(f"BDD {instance.name} — Transfert terminé", msg)
        _set_status(f"{BOX_NAME} — Arrêté")
        pythoncom.CoUninitialize()


# ── File watcher ──────────────────────────────────────────────────────────────

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


def _run_background(file_queue: queue.Queue) -> None:
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — watching for images.")
        return obs

    observer = _start_observer()
    inst = get_selected_instance()
    _set_status(f"{BOX_NAME} — BDD : {inst.name} — Prêt", processing=False)

    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observer has stopped (network drop?). Attempting reconnect...")
                _set_status(f"{BOX_NAME} — Reconnexion...", processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                log.info(f"Waiting {_RECONNECT_WAIT}s before restarting observer...")
                time.sleep(_RECONNECT_WAIT)
                observer = _start_observer()
                inst = get_selected_instance()
                _set_status(f"{BOX_NAME} — BDD : {inst.name} — Prêt", processing=False)
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


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    global _icon, _mutex_handle

    # Ensure only one instance of this router runs at a time.
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Multi_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

    # When launched by double-click from Explorer, refuse to start if SV is already running
    # to prevent conflicts with file handles and COM objects.
    try:
        parent_name = psutil.Process(os.getpid()).parent().name().lower()
    except Exception:
        parent_name = ""

    if parent_name == "explorer.exe":
        sv_running = any(
            (p.info["name"] or "").lower() in _ALL_SV_EXES
            for p in psutil.process_iter(["name"])
        )
        if sv_running:
            ctypes.windll.user32.MessageBoxW(
                0,
                "Pour relancer le routeur d'images, veuillez fermer complètement "
                "tous les StudioVision puis relancer.",
                "Routeur d'images",
                0x30,
            )
            sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("MULTI-INSTANCE version started (sélection manuelle de la BDD)")
    log.info(f"  Source      : {SOURCE_DIR}")
    log.info(f"  Orphelins   : {ORPHAN_DIR}")
    log.info(f"  BDD défaut  : {_DEFAULT_INSTANCE_NAME}")
    for inst in INSTANCES:
        log.info(f"  [{inst.name}]  exe={inst.exe}  mdb={inst.public_mdb}")
    log.info(f"  Timeout     : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Extensions  : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    ).start()

    threading.Thread(
        target=_run_background, args=(file_queue,), name="Background", daemon=True
    ).start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without system tray.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested.")
        finally:
            _stop_event.set()
        return

    # ── Build pystray menu ────────────────────────────────────────────────
    menu = pystray.Menu(
        # Non-clickable status line
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        # Radio-style instance selector
        pystray.MenuItem(
            "Base de données",
            _make_instance_menu(),
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
    log.info("Application stopped.")


if __name__ == "__main__":
    main()