"""
Routes incoming imaging files to the correct patient folder,
inserts a DB record, and refreshes the Access UI.

MULTI-INSTANCE VERSION: a single program monitors SOURCE_DIR and automatically
determines which StudioVision instance has the patient open (Megret or Romoli).
The instance with the active patient record receives the file.

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

DOCTOR_NAME_TO_INSTANCE: dict[str, str] = {
    "OM":  "Megret",
    "OMB": "Megret",
    "HR":  "Romoli",
    "HRB": "Romoli",
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


def patient_exists_in_db(patient_code: str, instance: Instance) -> bool:
    """Return True if the patient code exists in this instance's PUBLIC.MDB."""
    if not PYODBC_AVAILABLE:
        return False
    if not instance.public_mdb.exists():
        log.debug(f"[{instance.name}] PUBLIC.MDB not accessible: {instance.public_mdb}")
        return False
    try:
        conn   = db_connect(instance.public_mdb)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 [code patient] FROM Documents WHERE [code patient] = ?",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()
        return row is not None
    except Exception as e:
        log.debug(f"[{instance.name}] patient_exists_in_db error: {e}")
        return False


def find_instance_from_last_consult(patient_code: str) -> "Instance | None":
    if not PYODBC_AVAILABLE:
        return None

    @dataclass
    class ConsultResult:
        inst:     Instance
        date:     datetime
        dr_name:  str | None
        resolved: "Instance | None"  # instance déduite du Dr, ou None si neutre

    results: list[ConsultResult] = []

    for inst in INSTANCES:
        if not inst.public_mdb.exists():
            log.debug(f"[{inst.name}] PUBLIC.MDB inaccessible.")
            continue
        try:
            conn   = db_connect(inst.public_mdb)
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT TOP 1
                    c.[Date],
                    m.[Name]
                FROM [Consultations] AS c
                LEFT JOIN [Médecins] AS m
                    ON m.[ID] = c.[Code Médecin]
                WHERE c.[Code patient] = ?
                  AND c.[Code Médecin] IS NOT NULL
                ORDER BY c.[Date] DESC
                """,
                (int(patient_code),),
            )
            row = cursor.fetchone()
            conn.close()

            if not row or not row[0]:
                log.debug(f"[{inst.name}] Aucune consultation pour patient {patient_code}.")
                continue

            record_date = (
                row[0] if isinstance(row[0], datetime)
                else datetime.fromisoformat(str(row[0]))
            )
            dr_name  = str(row[1]).strip().upper() if row[1] else None
            target   = DOCTOR_NAME_TO_INSTANCE.get(dr_name) if dr_name else None
            resolved = next((i for i in INSTANCES if i.name == target), None) if target else None

            log.debug(
                f"[{inst.name}] Dernière consult : "
                f"Date={record_date:%Y-%m-%d}  Dr={dr_name}  → résolu={target}"
            )
            results.append(ConsultResult(inst, record_date, dr_name, resolved))

        except Exception as e:
            log.debug(f"[{inst.name}] Erreur lecture Consultations : {e}")

    if not results:
        log.warning(f"Aucune consultation trouvée pour patient {patient_code}.")
        return None

    # Trier par date décroissante — le record le plus récent est la source de vérité.
    results.sort(key=lambda r: r.date, reverse=True)
    best = results[0]

    if best.resolved is not None:
        log.info(
            f"Instance résolue via Dr '{best.dr_name}' "
            f"(date={best.date:%Y-%m-%d}, base=[{best.inst.name}]) "
            f"→ [{best.resolved.name}]"
        )
        return best.resolved

    # Dr neutre (PHF, LM, CS…) : on retourne l'instance dont la base
    # contient le record le plus récent — meilleure heuristique disponible.
    log.warning(
        f"Dr neutre '{best.dr_name}' pour patient {patient_code} "
        f"— fallback sur la base la plus récente : [{best.inst.name}] "
        f"(date={best.date:%Y-%m-%d})"
    )
    return best.inst


def find_instance_for_patient(patient_code: str) -> "Instance | None":
    """
    Détermine l'instance active pour un patient donné.

    Stratégie (ordre de priorité) :
      1. Lecture du Dr du dernier acte via jointure Médecins → le champ Name
         (OM/OMB → Megret, HR/HRB → Romoli) est la source de vérité.
         La base avec la consultation la plus récente l'emporte quand les
         deux bases contiennent le même patient (OM copie souvent HR).
      2. Fallback : recherche dans les bases des instances actives
         (comportement d'origine, conservé si la table Consultations est
         inaccessible ou ne contient aucune ligne pour ce patient).
    """
    # --- Stratégie 1 : Dr du dernier acte -----------------------------------
    inst = find_instance_from_last_consult(patient_code)
    if inst is not None:
        return inst

    log.warning(
        f"Stratégie Dr échouée pour le patient {patient_code} — "
        "repli sur la détection par processus/base."
    )

    # --- Stratégie 2 (fallback) : instances avec processus actif -----------
    running_exes = {
        (p.info["name"] or "").lower()
        for p in psutil.process_iter(["name"])
    }
    active_instances = [i for i in INSTANCES if i.exe.lower() in running_exes]
    if not active_instances:
        log.warning("Aucun processus StudioVision détecté — recherche dans toutes les bases.")
        active_instances = INSTANCES

    for inst in active_instances:
        if patient_exists_in_db(patient_code, inst):
            log.info(f"Patient {patient_code} trouvé dans l'instance [{inst.name}] (fallback).")
            return inst

    log.warning(f"Patient {patient_code} introuvable dans toutes les instances.")
    return None


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


def _instance_from_access_path(db_path: str) -> "Instance | None":
    """
    Given the CurrentDb path reported by an Access COM object,
    return the matching Instance by comparing exact folder names.
    """
    norm = db_path.lower().replace("/", "\\")
    for inst in INSTANCES:
        folder = inst.dest_photos.parent.name.lower()
        folder_pattern = f"\\{folder}\\"
        if folder_pattern in norm:
            return inst
    return None


def _instance_from_window_title(title: str) -> "Instance | None":
    title_lower = title.lower()
    sorted_instances = sorted(INSTANCES, key=lambda i: len(i.dest_photos.parent.name), reverse=True)

    for inst in sorted_instances:
        folder = inst.dest_photos.parent.name.lower()
        if folder in title_lower:
            return inst
    return None


def _get_access_pids() -> dict[int, str]:
    """
    Return {pid: window_title} for every visible Access top-level window.
    Uses EnumWindows so we can resolve which Access process owns which window.
    """
    import win32gui
    import win32process

    result: dict[int, str] = {}

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            return
        # Only keep Access windows (check via psutil later)
        result[hwnd] = (pid, title)

    win32gui.EnumWindows(_cb, None)

    # Filter to msaccess.exe PIDs
    access_pids: dict[int, str] = {}
    try:
        pid_to_name = {p.info["pid"]: (p.info["name"] or "").lower()
                       for p in psutil.process_iter(["pid", "name"])}
    except Exception:
        return {}

    for hwnd, (pid, title) in result.items():
        if "msaccess" in pid_to_name.get(pid, ""):
            access_pids[pid] = title  # last window per PID wins — usually the main one

    return access_pids  # {pid: window_title}


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


def _get_access_com_objects() -> list[tuple[object, int]]:
    """
    Return a list of (Access COM object, pid) for every running msaccess.exe.

    Technique: iterate psutil for msaccess PIDs, then use
    win32com.client.GetObject with the process handle trick via
    AccessibleObjectFromWindow on the main Access hwnd.
    Fallback: GetActiveObject for a single instance.
    """
    import win32gui
    import win32process
    import win32com.client

    results: list[tuple[object, int]] = []

    # Map Access window titles by PID so we can later resolve instance from window title.
    access_pid_titles: dict[int, str] = {}
    try:
        def _cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                access_pid_titles[pid] = title
            except Exception:
                pass
        win32gui.EnumWindows(_cb, None)
    except Exception as e:
        log.debug(f"EnumWindows failed: {e}")

    # Get all msaccess PIDs
    access_pids = [
        p.info["pid"]
        for p in psutil.process_iter(["pid", "name"])
        if "msaccess" in (p.info["name"] or "").lower()
    ]

    if not access_pids:
        return results

    if len(access_pids) == 1:
        # Single Access process — avoid the more complex multi-instance path.
        try:
            access = win32com.client.GetActiveObject("Access.Application")
            results.append((access, access_pids[0]))
        except Exception as e:
            log.debug(f"GetActiveObject failed: {e}")
        return results

    # Multiple instances: bind each Access window via its HWND using the
    # "OMain" window class, which is the top-level frame for every Access instance.
    OBJID_NATIVEOM = -16  # OBJID_NATIVEOM constant
    try:
        import ctypes
        import ctypes.wintypes

        acc_hwnds: list[tuple[int, int]] = []  # (hwnd, pid)

        def _find_access_main(hwnd, _):
            cls = win32gui.GetClassName(hwnd)
            if cls == "OMain":
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    acc_hwnds.append((hwnd, pid))
                except Exception:
                    pass

        win32gui.EnumWindows(_find_access_main, None)

        IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
        ole32 = ctypes.oledll.oleacc

        for hwnd, pid in acc_hwnds:
            try:
                ptr = ctypes.c_void_p()
                riid = pythoncom.MakeIID(IID_IDispatch)
                # AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, IID_IDispatch, &ptr)
                hr = ctypes.windll.oleacc.AccessibleObjectFromWindow(
                    hwnd,
                    ctypes.c_uint32(0xFFFFFFF0),  # OBJID_NATIVEOM
                    ctypes.byref(pythoncom.MakeIID(IID_IDispatch)
                                 if False else ctypes.create_string_buffer(16)),
                    ctypes.byref(ptr),
                )
                # Low-level OBJID_NATIVEOM binding is fragile across Office versions;
                # fall through to the simpler GetObject moniker approach below.
            except Exception:
                pass

        # Office supports a "!hwnd" moniker syntax that returns the document
        # COM object for a specific window handle, avoiding GetActiveObject
        # which always returns the same (last registered) instance.
        for hwnd, pid in acc_hwnds:
            try:
                access = win32com.client.GetObject(f"!{hwnd}")
                # GetObject with an hwnd moniker returns the document; navigate up to Application.
                try:
                    app = access.Application
                except Exception:
                    app = access
                results.append((app, pid))
                log.debug(f"Bound Access COM via hwnd {hwnd} (pid={pid})")
            except Exception as e:
                log.debug(f"GetObject(!{hwnd}) failed: {e}")

        if results:
            return results

    except Exception as e:
        log.debug(f"Multi-instance COM binding failed: {e}")

    # Last resort: GetActiveObject always returns the same registered instance,
    # so this only works correctly when a single Access process is running.
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        pid    = access_pids[0] if access_pids else 0
        results.append((access, pid))
    except Exception as e:
        log.debug(f"GetActiveObject fallback failed: {e}")

    return results


def get_active_patient() -> "dict | None":
    if not WIN32_AVAILABLE:
        return None

    import win32gui
    import win32process

    fg_access_pid: int | None = None
    try:
        fg_hwnd = win32gui.GetForegroundWindow()
        if fg_hwnd:
            _, fg_pid = win32process.GetWindowThreadProcessId(fg_hwnd)
            for p in psutil.process_iter(["pid", "name"]):
                if p.info["pid"] == fg_pid and "msaccess" in (p.info["name"] or "").lower():
                    fg_access_pid = fg_pid
                    break
    except Exception as e:
        log.debug(f"Foreground PID detection failed: {e}")

    log.debug(f"Foreground Access PID: {fg_access_pid}")

    com_objects = _get_access_com_objects()
    if not com_objects:
        log.debug("No Access COM objects found.")
        return None

    best_patient:  dict | None     = None
    best_instance: Instance | None = None
    best_is_fg:    bool            = False

    for access, pid in com_objects:
        # Identify instance from CurrentDb path
        inst: Instance | None = None
        try:
            db_path = access.CurrentDb().Name
            inst    = _instance_from_access_path(db_path)
            log.debug(f"Access pid={pid} db={db_path} → instance={inst.name if inst else 'unknown'}")
        except Exception as e:
            log.debug(f"CurrentDb() failed for pid={pid}: {e}")

        patient = _read_patient_from_access(access)
        if patient is None:
            log.debug(f"No active patient in Access pid={pid}")
            continue

        patient["instance"] = inst
        is_fg = (fg_access_pid is not None and pid == fg_access_pid)

        log.debug(
            f"Patient found: code={patient['code']} pid={pid} "
            f"instance={inst.name if inst else 'unknown'} fg={is_fg}"
        )

        if is_fg or best_patient is None:
            best_patient  = patient
            best_instance = inst
            best_is_fg    = is_fg
            if is_fg:
                break  # foreground match — no need to look further

    if best_patient is not None:
        log.info(
            f"Active patient: {best_patient['nom']} {best_patient['prenom']} "
            f"(code {best_patient['code']}) "
            f"instance={best_instance.name if best_instance else 'unknown'} "
            f"fg={best_is_fg}"
        )
    return best_patient


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
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh = False
                    last_patient_code = None
                    _notify("Transfer complete", f"{burst_count} file(s) processed")
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

            instance = patient.get("instance")
            if instance is not None:
                log.info(
                    f"Instance locked from active window: [{instance.name}] "
                    f"(patient {patient['code']})"
                )
            else:
                log.warning(
                    f"Could not lock instance from active window for patient "
                    f"{patient['code']} — falling back to stratégie Dr/Médecins."
                )
                instance = find_instance_for_patient(patient["code"])

            if instance is None:
                log.error(
                    f"Patient {patient['code']} not found in any instance. "
                    "Orphaning."
                )
                orphan_file(file)
                _notify("Orphan file", file.name)
                file_queue.task_done()
                continue

            patient_folder = find_patient_folder(patient["code"], instance)
            if not patient_folder:
                log.error(
                    f"[{instance.name}] Could not resolve folder for patient {patient['code']}. "
                    "Orphaning."
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
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description, instance):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")
                _notify("DB Error", "Insert failed — check logs")

            file_queue.task_done()

    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                _notify("Transfer complete", f"{burst_count} file(s) processed")
        _set_status(f"{BOX_NAME} — Stopped")
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
                log.warning("Observer has stopped (network drop?). Attempting reconnect...")
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

    log.info("MULTI-INSTANCE version started")
    log.info(f"  Source  : {SOURCE_DIR}")
    log.info(f"  Orphans : {ORPHAN_DIR}")
    for inst in INSTANCES:
        log.info(f"  [{inst.name}]  exe={inst.exe}  mdb={inst.public_mdb}")
    log.info(f"  Timeout : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext     : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info(f"  Dr→Instance : {DOCTOR_NAME_TO_INSTANCE}")

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

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open logs", _open_logs),
        pystray.MenuItem("Quit", _quit),
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