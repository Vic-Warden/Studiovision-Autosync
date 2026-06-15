"""
Fronto — Refractometer serial bridge for StudioVision
Reads frames from the refractometer on COM6 and injects values
into the REFRACTION form open in Access via win32com.
"""

from typing import Dict, Optional, Set
import win32com.client
import pythoncom
import serial
import threading
import subprocess
import sys
import re
import time
from pathlib import Path
import psutil
import logging

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

# Logging — writes to fronto.log and stdout
logging.basicConfig(
    level=logging.DEBUG,  # Use INFO in production
    format="%(asctime)s [%(levelname)-8s] %(threadName)-15s: %(message)s",
    handlers=[
        logging.FileHandler("fronto.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("Fronto")

# Configuration
APP_NAME    = "Fronto"
SERIAL_PORT = "COM6"
BAUD_RATE   = 9600
BYTESIZE    = serial.EIGHTBITS
PARITY      = serial.PARITY_NONE
STOPBITS    = serial.STOPBITS_ONE

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

_SV_POLL_INTERVAL   = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT = 30  # Seconds to wait for msaccess.exe to appear

# Global state
_stop_event  = threading.Event()
_icon        = None
_status_text = "Starting..."

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)
_COLOR_ERROR  = (220, 50, 50)


# System tray helpers
def _make_icon(color):
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse([margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin], fill=color)
    return img


def _set_status(text, color=None):
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            if color is not None:
                _icon.icon = _make_icon(color)
            _icon.update_menu()
        except Exception as e:
            logger.debug("Tray update failed: %s", e)


def _notify(title, message=""):
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            logger.debug("Notification failed: %s", e)


def _quit(icon, item):
    logger.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()


def parse_trame(line):
    # type: (str) -> Optional[Dict[str, str]]
    """
    Parses a refractometer serial frame.
    Examples:
      [01]RSPH=-11.25;CYL=-02.50;AXS=028;AD1=+01.75;Phx=+31.27;[17]
      [02]LSPH=-12.50;CYL=-01.00;AXS=148;AD1=+03.50;Phx=+32.46;IDP=43288;[17]
    Returns a dict with 'eye' ('OD' or 'OG') and parsed values, or None.
    """
    line = line.strip()
    logger.debug("Parsing frame: '%s'", line)

    # [01] = right eye (OD), [02] = left eye (OG)
    eye_match = re.match(r'^\[0([12])\]', line)
    if not eye_match:
        logger.warning("Unrecognised eye prefix — expected [01] or [02]. Skipping.")
        return None

    eye = "OD" if eye_match.group(1) == "1" else "OG"
    result = {"eye": eye}
    logger.debug("Eye detected: %s", eye)

    m = re.search(r'[RL]SPH=([+-]?\d+\.\d+)', line)
    if m:
        result["SPH"] = m.group(1)

    m = re.search(r'CYL=([+-]?\d+\.\d+)', line)
    if m:
        result["CYL"] = m.group(1)

    m = re.search(r'AXS=(\d+)', line)
    if m:
        result["AXS"] = str(int(m.group(1)))  # strip leading zeros

    m = re.search(r'AD1=([+-]?\d+\.\d+)', line)
    if m:
        result["ADD"] = m.group(1)

    if len(result) > 1:
        logger.info("Frame decoded: %s", result)
        return result
    else:
        logger.warning("No clinical values (SPH, CYL, AXS, ADD) found in frame.")
        return None


def inject_into_access(data):
    # type: (Dict[str, str]) -> None
    """Injects parsed refractometer values into the REFRACTION form in Access."""
    logger.info("Injecting into Access...")
    pythoncom.CoInitialize()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        logger.error("Could not connect to StudioVision via COM: %s", e)
        _set_status("{} — Access error".format(APP_NAME), _COLOR_ERROR)
        return

    try:
        form = access.Forms("REFRACTION")
    except Exception as e:
        logger.error("Form 'REFRACTION' not open in Access: %s", e)
        _set_status("{} — Form not found".format(APP_NAME), _COLOR_ERROR)
        return

    eye = data["eye"]
    mapping = {
        "SPH": "SPHERE {}".format(eye),
        "CYL": "CYLINDRE {}".format(eye),
        "AXS": "AXE {}".format(eye),
        "ADD": "ADD {}".format(eye),
    }

    success_count = 0
    for key, field_name in mapping.items():
        if key in data:
            try:
                form.Controls(field_name).Value = data[key]
                logger.info("  %s = %s", field_name, data[key])
                success_count += 1
            except Exception as e:
                logger.error("  Field '%s' write failed: %s", field_name, e)

    logger.info("Injection complete (%d field(s) written for %s).", success_count, eye)
    _set_status("{} — Data sent ({})".format(APP_NAME, eye), _COLOR_ACTIVE)
    _notify("Refractometer", "{} value(s) injected for {}".format(success_count, eye))
    time.sleep(2)
    _set_status("{} — Waiting".format(APP_NAME), _COLOR_READY)


def monitor_serial():
    """Continuously reads the serial port and dispatches parsed frames to Access."""
    logger.info("Opening %s at %d bps...", SERIAL_PORT, BAUD_RATE)
    while not _stop_event.is_set():
        try:
            with serial.Serial(
                port=SERIAL_PORT,
                baudrate=BAUD_RATE,
                bytesize=BYTESIZE,
                parity=PARITY,
                stopbits=STOPBITS,
                timeout=1,
            ) as ser:
                logger.info("%s open — waiting for frames...", SERIAL_PORT)
                _set_status("{} — Waiting".format(APP_NAME), _COLOR_READY)
                buffer = ""
                while not _stop_event.is_set():
                    raw = ser.read(256)
                    if raw:
                        buffer += raw.decode("ascii", errors="replace")
                        while "\n" in buffer:
                            line, buffer = buffer.split("\n", 1)
                            line = line.strip()
                            if not line:
                                continue
                            logger.info("Frame received: %s", line)
                            data = parse_trame(line)
                            if data:
                                # Inject in a separate thread to avoid blocking serial reads
                                t = threading.Thread(
                                    target=inject_into_access,
                                    args=(data,),
                                    name="Inject_{}".format(data["eye"]),
                                )
                                t.daemon = True
                                t.start()

        except serial.SerialException as e:
            logger.error("Serial port error on %s: %s — retrying in 5s...", SERIAL_PORT, e)
            _set_status("{} — Serial port error".format(APP_NAME), _COLOR_ERROR)
            time.sleep(5)
        except Exception as e:
            logger.critical("Unexpected error in monitor_serial: %s", e)
            _set_status("{} — Unexpected error".format(APP_NAME), _COLOR_ERROR)
            time.sleep(5)

    logger.info("Serial thread stopped.")


def _get_msaccess_pids():
    """Returns the set of all running msaccess.exe PIDs."""
    pids = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision():
    """
    Launches StudioVision and monitors msaccess.exe.
    Triggers shutdown when StudioVision is closed.
    Force-kills zombie msaccess.exe processes on exit to release COM locks.
    """
    logger.info("Launching StudioVision...")
    _set_status("{} — Launching StudioVision...".format(APP_NAME), _COLOR_READY)

    pids_before = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        logger.error("Executable not found: %s. Shutting down.", STUDIO_VISION_CMD[0])
        _set_status("{} — Executable not found".format(APP_NAME), _COLOR_ERROR)
        _stop_event.set()
        return
    except Exception as e:
        logger.error("Could not launch StudioVision: %s", e)
        _set_status("{} — Launch error".format(APP_NAME), _COLOR_ERROR)
        _stop_event.set()
        return

    logger.info("Waiting for msaccess.exe (max %ds)...", _SV_STARTUP_TIMEOUT)
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            logger.info("StudioVision is running.")
            _set_status("{} — StudioVision running".format(APP_NAME), _COLOR_READY)
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            logger.error("msaccess.exe did not appear. Shutting down.")
            _set_status("{} — SV did not start".format(APP_NAME), _COLOR_ERROR)
            _stop_event.set()
        return

    consecutive_empty = 0
    _EMPTY_THRESHOLD  = 2
    tracked_pids = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            current_pids = _get_msaccess_pids() - pids_before
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                logger.debug("StudioVision missing (%d/%d).", consecutive_empty, _EMPTY_THRESHOLD)
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    logger.info("StudioVision closed by user. Initiating shutdown.")
                    break
            else:
                consecutive_empty = 0
    except Exception as e:
        logger.error("Error monitoring msaccess.exe: %s", e)
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    logger.warning("Force-killed zombie msaccess.exe (PID %d).", pid)
            except Exception:
                pass

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass
        logger.info("Lifecycle thread stopped.")


def main():
    global _icon

    logger.info("Fronto starting.")

    sv_thread = threading.Thread(target=_launch_studio_vision, name="SV_Launcher")
    sv_thread.daemon = True
    sv_thread.start()

    serial_thread = threading.Thread(target=monitor_serial, name="SerialMonitor")
    serial_thread.daemon = True
    serial_thread.start()

    if not TRAY_AVAILABLE:
        logger.warning("pystray/Pillow not available — running headless.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("Keyboard interrupt.")
        finally:
            _stop_event.set()
        logger.info("Application stopped.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(text=lambda item: _status_text, action=None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", _quit),
    )

    _icon = pystray.Icon(
        name=APP_NAME,
        icon=_make_icon(_COLOR_READY),
        title=APP_NAME,
        menu=menu,
    )

    logger.info("System tray icon started.")
    _icon.run()

    _stop_event.set()
    sv_thread.join(timeout=15)
    serial_thread.join(timeout=5)
    logger.info("Application stopped.")


if __name__ == "__main__":
    main()