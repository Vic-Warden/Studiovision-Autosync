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

# Configuration
SERIAL_PORT  = "COM6"
BAUD_RATE    = 9600
BYTESIZE     = serial.EIGHTBITS
PARITY       = serial.PARITY_NONE
STOPBITS     = serial.STOPBITS_ONE

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

# Global stop event — set by _launch_studio_vision() or KeyboardInterrupt
_stop_event = threading.Event()  # type: threading.Event

# Studio Vision lifecycle constants
_SV_POLL_INTERVAL   = 3   # Seconds between msaccess.exe alive checks
_SV_STARTUP_TIMEOUT = 30  # Seconds to wait for msaccess.exe to appear after launch


def parse_trame(line):
    # type: (str) -> Optional[Dict[str, str]]
    """
    Parses a refractometer serial frame.
    Examples:
      [01]RSPH=-11.25;CYL=-02.50;AXS=028;AD1=+01.75;Phx=+31.27;[17]
      [02]LSPH=-12.50;CYL=-01.00;AXS=148;AD1=+03.50;Phx=+32.46;IDP=43288;[17]
    Returns a dict with 'eye' ('OD' or 'OG') and parsed values, or None if unrecognised.
    """
    line = line.strip()

    # [01] = right eye (OD), [02] = left eye (OG)
    eye_match = re.match(r'^\[0([12])\]', line)
    if not eye_match:
        return None
    eye = "OD" if eye_match.group(1) == "1" else "OG"

    result = {"eye": eye}  # type: Dict[str, str]

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

    return result if len(result) > 1 else None


def inject_into_access(data):
    # type: (Dict[str, str]) -> None
    """Injects parsed refractometer values into the REFRACTION form open in Access."""
    pythoncom.CoInitialize()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        print("[Access] Could not connect to StudioVision: {}".format(e))
        return

    try:
        form = access.Forms("REFRACTION")
    except Exception as e:
        print("[Access] Form REFRACTION not found: {}".format(e))
        return

    eye = data["eye"]

    mapping = {
        "SPH": "SPHERE {}".format(eye),
        "CYL": "CYLINDRE {}".format(eye),
        "AXS": "AXE {}".format(eye),
        "ADD": "ADD {}".format(eye),
    }  # type: Dict[str, str]

    for key, field_name in mapping.items():
        if key in data:
            try:
                form.Controls(field_name).Value = data[key]
                print("[Access] {:<15} = {}".format(field_name, data[key]))
            except Exception as e:
                print("[Access] {}: {}".format(field_name, e))


def monitor_serial():
    # type: () -> None
    """Continuously reads the serial port and dispatches parsed frames to Access."""
    print("[Serial] Opening {} at {} bps...".format(SERIAL_PORT, BAUD_RATE))
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
                print("[Serial] {} open — waiting for frames...".format(SERIAL_PORT))
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
                            print("[Serial] <- {}".format(line))
                            data = parse_trame(line)
                            if data:
                                # Inject in a separate thread to avoid blocking serial reads
                                t = threading.Thread(
                                    target=inject_into_access,
                                    args=(data,),
                                )
                                t.daemon = True
                                t.start()

        except serial.SerialException as e:
            print("[Serial] Error: {} — retrying in 5s...".format(e))
            time.sleep(5)
        except Exception as e:
            print("[Serial] Unexpected error: {}".format(e))
            time.sleep(5)

    print("[Serial] Thread stopped.")


def _get_msaccess_pids():
    # type: () -> Set[int]
    """Returns the set of all running msaccess.exe PIDs."""
    pids = set()  # type: Set[int]
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision():
    # type: () -> None
    """
    Launches StudioVision and monitors msaccess.exe.
    When StudioVision is closed by the user, kills zombie processes and
    signals the rest of the application to stop via _stop_event.
    """
    print("[SV] Launching StudioVision: {}".format(" ".join(STUDIO_VISION_CMD)))

    pids_before = _get_msaccess_pids()  # type: Set[int]

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        print("[SV] Executable not found. Shutting down.")
        _stop_event.set()
        return
    except Exception as e:
        print("[SV] Could not launch StudioVision: {}. Shutting down.".format(e))
        _stop_event.set()
        return

    print("[SV] Waiting up to {}s for msaccess.exe to start...".format(_SV_STARTUP_TIMEOUT))
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            print("[SV] StudioVision is starting...")
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            print("[SV] msaccess.exe did not appear. Shutting down.")
            _stop_event.set()
        return

    consecutive_empty = 0
    _EMPTY_THRESHOLD  = 2
    tracked_pids = set()  # type: Set[int]

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)

            current_pids = _get_msaccess_pids() - pids_before
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                print("[SV] StudioVision missing ({}/{}).".format(consecutive_empty, _EMPTY_THRESHOLD))
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    print("[SV] StudioVision closed by user. Initiating shutdown.")
                    break
            else:
                consecutive_empty = 0

    except Exception as e:
        print("[SV] Error while monitoring msaccess.exe: {}".format(e))
    finally:
        # Kill any zombie msaccess.exe processes to release COM locks
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    print("[SV] Force-killed zombie msaccess.exe (PID {}).".format(pid))
            except Exception:
                pass

        # Wait for all killed PIDs to fully disappear before signalling stop
        _KILL_DRAIN_TIMEOUT = 10
        _KILL_DRAIN_POLL    = 0.5
        deadline = time.monotonic() + _KILL_DRAIN_TIMEOUT
        while time.monotonic() < deadline:
            still_alive = {pid for pid in tracked_pids if psutil.pid_exists(pid)}
            if not still_alive:
                break
            print("[SV] Waiting for PIDs to fully exit: {}".format(still_alive))
            time.sleep(_KILL_DRAIN_POLL)

        _stop_event.set()
        print("[SV] StudioVision lifecycle thread stopped.")


if __name__ == "__main__":
    # Launch StudioVision in a dedicated thread (monitors process and drives shutdown)
    sv_thread = threading.Thread(target=_launch_studio_vision, name="StudioVisionLauncher")
    sv_thread.daemon = True
    sv_thread.start()

    # Run serial monitoring in a dedicated thread so KeyboardInterrupt is catchable here
    serial_thread = threading.Thread(target=monitor_serial, name="SerialMonitor")
    serial_thread.daemon = True
    serial_thread.start()

    try:
        # Main thread just waits for the stop signal
        while not _stop_event.is_set():
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[Main] Keyboard interrupt — shutting down...")
        _stop_event.set()

    sv_thread.join(timeout=15)
    serial_thread.join(timeout=5)
    print("[Main] Application stopped.")