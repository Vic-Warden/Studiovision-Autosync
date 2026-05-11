"""
windows7V4.py
Image Router — Version 4, Windows 7 compatible build.

Combines all improvements introduced in Version 4 (PollingObserver, network
share wait, auto-reconnect, targeted SFDoc Requery, dirty-state guard, Requery
retry loop, patient-code guard, log in ~/studiovision/, system tray icon) with
Windows 7 / Python 3.9 compatibility constraints (typing.Optional instead of
X | Y union syntax, str.format() where needed, daemon=True set as an
attribute).

Dependencies: watchdog, pyodbc, pywin32, pythoncom, pystray, Pillow
"""

import os
import pythoncom
import queue
import shutil
import sys
import threading
import time
import logging
import ctypes
from datetime import datetime
from pathlib import Path
from typing import Optional
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

#  Configuration

BOX_NAME    = "Windows 7"

SOURCE_DIR  = Path(r"??")
ORPHAN_DIR  = Path(r"??")
DEST_PHOTOS = Path(r"??")
PUBLIC_MDB  = Path(r"??")
DOCUM_MDB   = Path(r"??")

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

# Log in ~/studiovision/ — valid as script and as compiled .exe
_LOG_DIR  = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, "image_router.log")

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

#  System Tray

_ICON_SIZE     = 64
_COLOR_READY   = (30, 144, 255)   # dodger blue — idle
_COLOR_ACTIVE  = (50, 205, 50)    # lime green  — processing

_icon          = None   # type: Optional[pystray.Icon]
_status_text   = "Starting..."  # type: str
_stop_event    = threading.Event()


def _make_icon(color):
    # type: (tuple) -> Image.Image
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse(
        [margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin],
        fill=color,
    )
    return img


def _set_status(text, processing=False):
    # type: (str, bool) -> None
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as e:
            log.debug("Tray update failed: %s", e)


def _notify(title, message=""):
    # type: (str, str) -> None
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            log.debug("Notification failed: %s", e)


def _open_logs(icon, item):
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning("Could not open log file: %s", e)


def _quit(icon, item):
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


#  Business logic (Python 3.9 compatible syntax)

def db_connect(mdb_path):
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(mdb_path) + ";"
    )


def get_active_patient():
    # type: () -> Optional[dict]
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data   = {}

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
        log.debug("COM error: %s", e)
        return None


def find_patient_folder(patient_code):
    # type: (str) -> Optional[Path]
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
            (int(patient_code),)
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
    except Exception as e:
        log.error("DB folder lookup failed: %s", e)
        return None


def insert_document(patient, relative_path, description):
    # type: (dict, str, str) -> bool
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
        log.info(
            "Insert OK: patient=%s path='%s' db=%s",
            patient["code"], relative_path, PUBLIC_MDB.name
        )
        return True
    except Exception as e:
        log.error("DB insert failed: %s", e)
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


def refresh_ui(expected_patient_code=None):
    # type: (Optional[str]) -> None
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
                    "Refresh skipped: expected patient %s but current patient is %s.",
                    expected_patient_code, current_code
                )
                return

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                "Subform '%s' not found in the active form. Refresh skipped.",
                SFDOC_SUBFORM_NAME
            )
            return

        try:
            form.Refresh()
            log.info("Refresh() on parent form (image repaint).")
        except Exception as e_pref:
            log.warning("Parent form Refresh() failed (non-blocking): %s", e_pref)

        try:
            if form.Dirty:
                log.info("Parent form is dirty — clearing edit mode before Requery().")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug("Could not read/clear form.Dirty: %s", e_dirty)

        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5
        requery_ok = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(
                    "Requery() on '%s' (attempt %d/%d).",
                    SFDOC_SUBFORM_NAME, attempt, _REQUERY_ATTEMPTS
                )
                requery_ok = True
                break
            except Exception as e_req:
                log.warning(
                    "Requery() attempt %d/%d failed on '%s': %s",
                    attempt, _REQUERY_ATTEMPTS, SFDOC_SUBFORM_NAME, e_req
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)

        if not requery_ok:
            log.warning(
                "All Requery() attempts failed on '%s', falling back to Refresh().",
                SFDOC_SUBFORM_NAME
            )
            try:
                sfdoc.Refresh()
                log.info("Fallback Refresh() on '%s'.", SFDOC_SUBFORM_NAME)
            except Exception as e_ref:
                log.warning(
                    "Fallback Refresh() also unavailable on '%s': %s",
                    SFDOC_SUBFORM_NAME, e_ref
                )

        try:
            sfdoc.Recordset.MoveLast()
            log.info("MoveLast() on '%s'", SFDOC_SUBFORM_NAME)
        except Exception as e_ml:
            log.debug("MoveLast() failed on '%s': %s", SFDOC_SUBFORM_NAME, e_ml)

    except Exception as e:
        log.warning("COM refresh failed (non-blocking): %s", e)


def wait_for_file(file):
    # type: (Path) -> bool
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug("File locked (%d/%d), retrying...", attempt, FILE_LOCK_MAX_ATTEMPTS)
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error("File still locked after %d attempts: %s", FILE_LOCK_MAX_ATTEMPTS, file)
    return False


def move_file(source, dest_folder, label=""):
    # type: (Path, Path, str) -> Optional[Path]
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / "{0}_{1}{2}".format(source.stem, ts, source.suffix)
        log.info("Name conflict, renamed to %s", dest.name)
    try:
        shutil.move(str(source), str(dest))
        tag = "[{0}]  ".format(label) if label else ""
        log.info("%s%s -> %s", tag, source.name, dest)
        return dest
    except Exception as e:
        log.error("Move failed: %s", e)
        return None


def orphan_file(file):
    # type: (Path) -> None
    log.warning("Orphaning: %s", file.name)
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def wait_for_network_share():
    # type: () -> None
    source_str = str(SOURCE_DIR)
    is_unc    = source_str.startswith("\\\\") or source_str.startswith("//")
    is_local  = (
        not is_unc
        and (len(source_str) >= 2 and source_str[1] == ":")
    )
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


def prevent_sleep():
    # type: () -> None
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning("Could not set execution state: %s", e)


#  Worker thread  (consumer)

def worker(file_queue):
    # type: (queue.Queue) -> None
    pythoncom.CoInitialize()
    log.info("Worker started.")

    needs_refresh     = False
    last_patient_code = None  # type: Optional[str]
    burst_count       = 0     # files successfully inserted in the current burst

    try:
        while True:
            try:
                file = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh     = False
                    last_patient_code = None
                    _notify(
                        "Transfer complete",
                        "{0} file(s) processed".format(burst_count),
                    )
                    _set_status("{0} — Ready".format(BOX_NAME), processing=False)
                    burst_count = 0
                continue
            except Exception as e:
                log.error("Queue error: %s", e)
                continue

            log.info("Processing: %s (%d pending)", file.name, file_queue.qsize())

            if burst_count == 0 and not needs_refresh:
                _notify("Transfer in progress", file.name)
            _set_status("Transfer in progress...", processing=True)

            if not file.exists():
                log.warning("File gone before processing: %s", file)
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error("Aborting, persistent lock: %s", file.name)
                _notify("Error", "File locked: {0}".format(file.name))
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
                        "No patient open, waiting (timeout in %d min)",
                        PATIENT_WAIT_TIMEOUT // 60
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                "Patient: %s %s (code %s)",
                patient["nom"], patient["prenom"], patient["code"]
            )

            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(
                    "Could not resolve folder for patient %s. Orphaning.",
                    patient["code"]
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
            relative_path = "\\{0}\\{1}\\{2}".format(
                group_name, patient_folder.name, dest.name
            )
            description = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            if insert_document(patient, relative_path, description):
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
                _notify(
                    "Transfer complete",
                    "{0} file(s) processed".format(burst_count),
                )
        _set_status("{0} — Stopped".format(BOX_NAME))
        pythoncom.CoUninitialize()


#  Watchdog producer

class ImageProducer(FileSystemEventHandler):
    def __init__(self, file_queue):
        # type: (queue.Queue) -> None
        super(ImageProducer, self).__init__()
        self._queue = file_queue

    def on_created(self, event):
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info("Enqueued: %s (queue size: %d)", file.name, self._queue.qsize() + 1)
        self._queue.put(file)


#  Background thread (observer + auto-reconnect loop)

def _run_background(file_queue):
    # type: (queue.Queue) -> None
    _RECONNECT_DELAY = 15

    def _start_observer():
        # type: () -> Observer
        producer = ImageProducer(file_queue)
        obs = Observer()
        obs.schedule(producer, str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started. Watching for images.")
        return obs

    observer = _start_observer()
    _set_status("{0} — Ready".format(BOX_NAME), processing=False)

    try:
        while not _stop_event.is_set():
            time.sleep(1)
            if not observer.is_alive():
                log.warning(
                    "Observer died (possible network drop). "
                    "Waiting %d s before reconnecting...",
                    _RECONNECT_DELAY
                )
                _set_status("{0} — Reconnecting...".format(BOX_NAME), processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                time.sleep(_RECONNECT_DELAY)
                log.info("Restarting Observer after network recovery.")
                observer = _start_observer()
                _set_status("{0} — Ready".format(BOX_NAME), processing=False)
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


#  Entry point

def main():
    # type: () -> None
    global _icon

    prevent_sleep()
    log.info("Checking network share availability...")
    wait_for_network_share()

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 4 started")
    log.info("  Source     : %s", SOURCE_DIR)
    log.info("  Dest       : %s", DEST_PHOTOS)
    log.info("  PUBLIC.MDB : %s", PUBLIC_MDB)
    log.info("  DOCUM.MDB  : %s", DOCUM_MDB)
    log.info("  Orphans    : %s", ORPHAN_DIR)
    log.info("  Timeout    : %d min", PATIENT_WAIT_TIMEOUT // 60)
    log.info("  Ext        : %s", ", ".join(sorted(WATCHED_EXTENSIONS)))

    file_queue = queue.Queue()

    worker_thread = threading.Thread(
        target=worker, args=(file_queue,), name="Worker"
    )
    worker_thread.daemon = True
    worker_thread.start()

    bg_thread = threading.Thread(
        target=_run_background, args=(file_queue,), name="Background"
    )
    bg_thread.daemon = True
    bg_thread.start()

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