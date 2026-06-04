import win32com.client
import pythoncom
import serial
import threading
import subprocess
import sys
import re
import time
from pathlib import Path

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


def parse_trame(line: str) -> dict | None:
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

    result = {"eye": eye}

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


def inject_into_access(data: dict):
    """Injects parsed refractometer values into the REFRACTION form open in Access."""
    pythoncom.CoInitialize()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        print(f"[Access] Could not connect to StudioVision: {e}")
        return

    try:
        form = access.Forms("REFRACTION")
    except Exception as e:
        print(f"[Access] Form REFRACTION not found: {e}")
        return

    eye = data["eye"]

    mapping = {
        "SPH": f"SPHERE {eye}",
        "CYL": f"CYLINDRE {eye}",
        "AXS": f"AXE {eye}",
        "ADD": f"ADD {eye}",
    }

    for key, field_name in mapping.items():
        if key in data:
            try:
                form.Controls(field_name).Value = data[key]
                print(f"[Access] {field_name:<15} = {data[key]}")
            except Exception as e:
                print(f"[Access] {field_name}: {e}")


def monitor_serial():
    """Continuously reads the serial port and dispatches parsed frames to Access."""
    print(f"[Serial] Opening {SERIAL_PORT} at {BAUD_RATE} bps...")
    while True:
        try:
            with serial.Serial(
                port=SERIAL_PORT,
                baudrate=BAUD_RATE,
                bytesize=BYTESIZE,
                parity=PARITY,
                stopbits=STOPBITS,
                timeout=1,
            ) as ser:
                print(f"[Serial] {SERIAL_PORT} open — waiting for frames...")
                buffer = ""
                while True:
                    raw = ser.read(256)
                    if raw:
                        buffer += raw.decode("ascii", errors="replace")
                        while "\n" in buffer:
                            line, buffer = buffer.split("\n", 1)
                            line = line.strip()
                            if not line:
                                continue
                            print(f"[Serial] <- {line}")
                            data = parse_trame(line)
                            if data:
                                # Inject in a separate thread to avoid blocking serial reads
                                t = threading.Thread(
                                    target=inject_into_access,
                                    args=(data,),
                                    daemon=True,
                                )
                                t.start()

        except serial.SerialException as e:
            print(f"[Serial] Error: {e} — retrying in 5s...")
            time.sleep(5)
        except Exception as e:
            print(f"[Serial] Unexpected error: {e}")
            time.sleep(5)


def launch_studio_vision():
    """Launches StudioVision and waits for Access to initialise."""
    exe = Path(STUDIO_VISION_CMD[0])
    if not exe.exists():
        print(f"[SV] Executable not found: {exe}")
        print("[SV] Attempting launch anyway...")
    print("[SV] Launching StudioVision...")
    try:
        subprocess.Popen(STUDIO_VISION_CMD)
        print("[SV] StudioVision launched.")
        time.sleep(8)  # Wait for Access to start
    except Exception as e:
        print(f"[SV] Could not launch StudioVision: {e}")
        sys.exit(1)


if __name__ == "__main__":
    launch_studio_vision()
    monitor_serial()