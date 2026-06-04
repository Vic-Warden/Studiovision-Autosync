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

# ==========================================
# CONFIGURATION DES LOGS
# ==========================================
# Crée un fichier fronto.log dans le dossier courant et affiche aussi dans la console
logging.basicConfig(
    level=logging.DEBUG, # DEBUG permet de tout voir. Mettre INFO en production.
    format="%(asctime)s [%(levelname)-8s] %(threadName)-15s: %(message)s",
    handlers=[
        logging.FileHandler("fronto.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("Fronto")

# Configuration globale
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

_SV_POLL_INTERVAL   = 3
_SV_STARTUP_TIMEOUT = 30

_stop_event   = threading.Event()
_icon         = None
_status_text  = "Starting..."

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)
_COLOR_ERROR  = (220, 50, 50)

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
            logger.debug(f"Erreur update icône: {e}")

def _notify(title, message=""):
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            logger.debug(f"Erreur notification: {e}")

def _quit(icon, item):
    logger.info("Fermeture demandée via l'icône de la barre des tâches.")
    _stop_event.set()
    icon.stop()

def parse_trame(line):
    # type: (str) -> Optional[Dict[str, str]]
    line = line.strip()
    logger.debug(f"Analyse de la ligne brute : '{line}'")

    eye_match = re.match(r'^\[0([12])\]', line)
    if not eye_match:
        logger.warning("-> Format d'œil non reconnu (ne commence pas par [01] ou [02]). Ignoré.")
        return None
    
    eye = "OD" if eye_match.group(1) == "1" else "OG"
    result = {"eye": eye}
    logger.debug(f"-> Œil détecté : {eye}")

    m = re.search(r'[RL]SPH=([+-]?\d+\.\d+)', line)
    if m:
        result["SPH"] = m.group(1)
        logger.debug(f"-> Sphère trouvée : {result['SPH']}")

    m = re.search(r'CYL=([+-]?\d+\.\d+)', line)
    if m:
        result["CYL"] = m.group(1)
        logger.debug(f"-> Cylindre trouvé : {result['CYL']}")

    m = re.search(r'AXS=(\d+)', line)
    if m:
        result["AXS"] = str(int(m.group(1)))
        logger.debug(f"-> Axe trouvé : {result['AXS']}")

    m = re.search(r'AD1=([+-]?\d+\.\d+)', line)
    if m:
        result["ADD"] = m.group(1)
        logger.debug(f"-> Addition trouvée : {result['ADD']}")

    if len(result) > 1:
        logger.info(f"Trame décodée avec succès : {result}")
        return result
    else:
        logger.warning("-> Aucune valeur clinique (SPH, CYL, AXS, ADD) trouvée dans la ligne.")
        return None

def inject_into_access(data):
    # type: (Dict[str, str]) -> None
    logger.info("Début de l'injection COM dans Access...")
    pythoncom.CoInitialize()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        logger.error(f"Impossible de se connecter à StudioVision via COM: {e}")
        _set_status(f"{APP_NAME} — Access error", _COLOR_ERROR)
        return

    try:
        form = access.Forms("REFRACTION")
    except Exception as e:
        logger.error(f"Le formulaire 'REFRACTION' n'est pas ouvert dans Access: {e}")
        _set_status(f"{APP_NAME} — Form not found", _COLOR_ERROR)
        return

    eye = data["eye"]
    mapping = {
        "SPH": f"SPHERE {eye}",
        "CYL": f"CYLINDRE {eye}",
        "AXS": f"AXE {eye}",
        "ADD": f"ADD {eye}",
    }

    succes_count = 0
    for key, field_name in mapping.items():
        if key in data:
            try:
                form.Controls(field_name).Value = data[key]
                logger.info(f"  ✅ Injection OK: {field_name:<15} = {data[key]}")
                succes_count += 1
            except Exception as e:
                logger.error(f"  ❌ Échec injection sur {field_name}: {e}")

    logger.info(f"Injection terminée ({succes_count} champs modifiés pour {eye}).")
    _set_status(f"{APP_NAME} — Data sent ({eye})", _COLOR_ACTIVE)
    _notify("Refractomètre", f"{succes_count} valeurs {eye} injectées")
    time.sleep(2)
    _set_status(f"{APP_NAME} — Waiting", _COLOR_READY)

def monitor_serial():
    logger.info(f"Tentative d'ouverture de {SERIAL_PORT} à {BAUD_RATE} bps...")
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
                logger.info(f"Port série {SERIAL_PORT} ouvert. En attente de données...")
                _set_status(f"{APP_NAME} — Waiting", _COLOR_READY)
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
                            logger.info(f"==> Trame reçue du port série : {line}")
                            data = parse_trame(line)
                            if data:
                                t = threading.Thread(target=inject_into_access, args=(data,), name=f"Inject_{data['eye']}")
                                t.daemon = True
                                t.start()

        except serial.SerialException as e:
            logger.error(f"Erreur Port Série {SERIAL_PORT}: {e}. Nouvelle tentative dans 5s...")
            _set_status(f"{APP_NAME} — Serial port error", _COLOR_ERROR)
            time.sleep(5)
        except Exception as e:
            logger.critical(f"Erreur inattendue dans monitor_serial: {e}")
            _set_status(f"{APP_NAME} — Unexpected error", _COLOR_ERROR)
            time.sleep(5)

    logger.info("Thread série arrêté.")

def _get_msaccess_pids():
    pids = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids

def _launch_studio_vision():
    logger.info("Démarrage de StudioVision...")
    _set_status(f"{APP_NAME} — Launching StudioVision...", _COLOR_READY)

    pids_before = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        logger.error(f"Exécutable introuvable: {STUDIO_VISION_CMD[0]}. Arrêt.")
        _set_status(f"{APP_NAME} — Executable not found", _COLOR_ERROR)
        _stop_event.set()
        return
    except Exception as e:
        logger.error(f"Impossible de lancer StudioVision: {e}")
        _set_status(f"{APP_NAME} — Launch error", _COLOR_ERROR)
        _stop_event.set()
        return

    logger.info(f"En attente du processus msaccess.exe (max {_SV_STARTUP_TIMEOUT}s)...")
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            logger.info("StudioVision est démarré.")
            _set_status(f"{APP_NAME} — StudioVision running", _COLOR_READY)
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            logger.error("msaccess.exe n'est pas apparu dans les temps. Arrêt.")
            _set_status(f"{APP_NAME} — SV did not start", _COLOR_ERROR)
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
                logger.debug(f"StudioVision absent ({consecutive_empty}/{_EMPTY_THRESHOLD}).")
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    logger.info("Fermeture de StudioVision détectée. Arrêt du pont.")
                    break
            else:
                consecutive_empty = 0
    except Exception as e:
        logger.error(f"Erreur lors de la surveillance de msaccess.exe: {e}")
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    logger.warning(f"Processus msaccess.exe (PID {pid}) tué de force pour libérer le COM.")
            except Exception:
                pass

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass
        logger.info("Cycle de vie SV arrêté.")

def main():
    global _icon

    logger.info("=== Démarrage de Fronto ===")
    
    sv_thread = threading.Thread(target=_launch_studio_vision, name="SV_Launcher")
    sv_thread.daemon = True
    sv_thread.start()

    serial_thread = threading.Thread(target=monitor_serial, name="SerialMonitor")
    serial_thread.daemon = True
    serial_thread.start()

    if not TRAY_AVAILABLE:
        logger.warning("Librairies systray (pystray/Pillow) non disponibles — exécution en mode console.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("Interruption clavier détectée.")
        finally:
            _stop_event.set()
        logger.info("Application arrêtée.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(text=lambda item: _status_text, action=None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quitter", _quit),
    )

    _icon = pystray.Icon(
        name=APP_NAME,
        icon=_make_icon(_COLOR_READY),
        title=APP_NAME,
        menu=menu,
    )

    logger.info("Icône système démarrée.")
    _icon.run()

    _stop_event.set()
    sv_thread.join(timeout=15)
    serial_thread.join(timeout=5)
    logger.info("=== Arrêt complet de Fronto ===")

if __name__ == "__main__":
    main()