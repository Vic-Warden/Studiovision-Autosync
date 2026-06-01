"""
Router/Dispatcher for Pierre Henri (Box-6).
Monitors the "Export CV" folder and moves images to "OM" or "HR"
based on the active systray selection.
"""

import os
import shutil
import time
import queue
import threading
import logging
import sys
import ctypes
from pathlib import Path

from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw

# Windows single-instance mutex
import win32api
import win32event
import winerror

# Configuration
SOURCE_DIR = Path(r"C:\Users\Box-6\Desktop\Export CV")
DEST_OM    = Path(r"C:\Users\Box-6\Desktop\OM")
DEST_HR    = Path(r"C:\Users\Box-6\Desktop\HR")

WATCHED_EXTENSIONS = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff", ".dcm", ".pdf"}

for d in [SOURCE_DIR, DEST_OM, DEST_HR]:
    d.mkdir(parents=True, exist_ok=True)

# State
_state_lock = threading.Lock()
_current_target = "OM"  # Default on startup
_stop_event = threading.Event()
_icon = None
_mutex_handle = None

COLOR_OM = (30, 144, 255)  # Blue
COLOR_HR = (50, 205, 50)   # Green

# Logging
log_file = Path(os.path.expanduser("~")) / "studiovision" / "triage.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler()]
)

# File routing
def wait_for_file(file_path: Path, retries=10, delay=1) -> bool:
    for _ in range(retries):
        try:
            with file_path.open("rb"):
                return True
        except (PermissionError, OSError):
            time.sleep(delay)
    return False

def worker_triage(file_queue: queue.Queue):
    logging.info("Triage worker started.")
    while not _stop_event.is_set():
        try:
            file_path = file_queue.get(timeout=1)
        except queue.Empty:
            continue

        if not file_path.exists() or not wait_for_file(file_path):
            file_queue.task_done()
            continue

        with _state_lock:
            target_name = _current_target
            dest_folder = DEST_OM if target_name == "OM" else DEST_HR

        dest_file = dest_folder / file_path.name

        if dest_file.exists():
            ts = int(time.time())
            dest_file = dest_folder / f"{file_path.stem}_{ts}{file_path.suffix}"

        try:
            shutil.move(str(file_path), str(dest_file))
            logging.info(f"OK: {file_path.name} -> {target_name}")
            if _icon:
                _icon.notify(f"Transferred to {target_name}", file_path.name)
        except Exception as e:
            logging.error(f"Move error for {file_path.name}: {e}")

        file_queue.task_done()

# Watchdog producer
class SourceHandler(FileSystemEventHandler):
    def __init__(self, q: queue.Queue):
        self.q = q

    def on_created(self, event):
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        if file_path.suffix.lower() in WATCHED_EXTENSIONS:
            logging.info(f"New file detected: {file_path.name}")
            self.q.put(file_path)

# System tray
def create_image(color):
    size = 64
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((4, 4, size - 4, size - 4), fill=color)
    return image

def set_target(icon, item):
    global _current_target
    with _state_lock:
        if "OM" in item.text:
            _current_target = "OM"
            icon.icon = create_image(COLOR_OM)
        else:
            _current_target = "HR"
            icon.icon = create_image(COLOR_HR)
    logging.info(f"Mode changed: routing to {_current_target}")
    icon.notify("Destination changed", f"Next images will go to {_current_target}")

def is_checked(target):
    def check(item):
        with _state_lock:
            return _current_target == target
    return check

def quit_app(icon, item):
    _stop_event.set()
    icon.stop()

def open_source_folder(icon, item):
    os.startfile(str(SOURCE_DIR))

# Entry point
def main():
    global _icon, _mutex_handle

    # Single-instance guard
    _mutex_handle = win32event.CreateMutex(None, False, "StudioVision_Export_Triage_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        sys.exit(0)

    q = queue.Queue()
    threading.Thread(target=worker_triage, args=(q,), daemon=True).start()

    observer = Observer()
    observer.schedule(SourceHandler(q), str(SOURCE_DIR), recursive=False)
    observer.start()
    logging.info(f"Monitoring started: {SOURCE_DIR}")

    menu = pystray.Menu(
        pystray.MenuItem("Open export folder", open_source_folder),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("-> Send to OM folder (Blue)", set_target, radio=True, checked=is_checked("OM")),
        pystray.MenuItem("-> Send to HR folder (Green)", set_target, radio=True, checked=is_checked("HR")),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", quit_app)
    )

    _icon = pystray.Icon("VisualFieldTriage", create_image(COLOR_OM), "Image Router", menu=menu)

    def on_ready(icon):
        icon.visible = True
        icon.notify(
            "Right-click the tray icon to select the destination folder (OM or HR).",
            "Image Router started"
        )

    _icon.run(setup=on_ready)

    observer.stop()
    observer.join()

if __name__ == "__main__":
    main()