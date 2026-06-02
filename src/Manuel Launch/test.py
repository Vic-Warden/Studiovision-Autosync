import os
import time
import shutil
import queue
import threading
import logging
import sys
from datetime import datetime
from pathlib import Path

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("Error: pywin32 not found.")
    sys.exit(1)

from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

# Configuration (adjust per workstation)
DOSSIER_SOURCE    = Path(r"C:\Users\Box-6\Desktop\Salle 6")
DOSSIER_ORPHELINS = Path(r"C:\Users\Box-6\Desktop\Images_Oubliees")
DEST_PHOTOS       = Path(r"\\studiovision\Studiov2000-OM\PHOTOS")

WATCHED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".dcm", ".pdf"}
_AC_SUBFORM = 112
SFDOC_SUBFORM_NAME = "SFDoc"

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("Routeur_V6")

def _find_sfdoc(form):
    """Recursively searches for the SFDoc subform."""
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

def calculer_dossier_studiovision(code: str, nom: str, prenom: str) -> str:
    """Returns the relative folder path using the Studio Vision naming convention."""
    code_str = str(code).strip()
    prefixe = code_str[:2].ljust(2, '0') + ".000"

    nom_clean = "".join(c for c in str(nom) if c.isalpha()).lower()
    prenom_clean = "".join(c for c in str(prenom) if c.isalpha()).lower()

    part1 = nom_clean[:3].ljust(3, 'a')
    part2 = prenom_clean[:3].ljust(3, 'a')

    return f"{prefixe}\\{code_str}{part1}.{part2}"

def move_file_to_patient(source: Path, nom_dossier_relatif: str) -> Path | None:
    """Moves the file to the correct patient folder on the network share."""
    dossier_patient = DEST_PHOTOS / nom_dossier_relatif
    dossier_patient.mkdir(parents=True, exist_ok=True)

    dest = dossier_patient / source.name
    if dest.exists():
        dest = dossier_patient / f"{source.stem}_{int(time.time())}{source.suffix}"

    try:
        shutil.move(str(source), str(dest))
        log.info(f"File moved to network: {dest}")
        return dest
    except Exception as e:
        log.error(f"Failed to move file to {dest}: {e}")
        return None

def worker(file_queue: queue.Queue):
    pythoncom.CoInitialize()
    log.info("Worker ready. Waiting for images...")

    while True:
        try:
            fichier: Path = file_queue.get()
        except Exception:
            continue

        if not fichier.exists():
            file_queue.task_done()
            continue

        log.info(f"New image detected: {fichier.name}")
        time.sleep(1)  # Allow the device to finish writing the file

        try:
            # Connect to the active Access interface
            access = win32com.client.GetActiveObject("Access.Application")
            form = access.Screen.ActiveForm

            if form is None:
                log.warning("No patient open. Moving file to orphan folder.")
                shutil.move(str(fichier), str(DOSSIER_ORPHELINS / fichier.name))
                file_queue.task_done()
                continue

            # Read patient info from the active form
            code_patient = str(form.Controls("Code patient").Value)
            nom_patient = str(form.Controls("NOM").Value)
            prenom_patient = str(form.Controls("Prénom").Value)

            log.info(f"Target patient: {nom_patient} {prenom_patient} (Code: {code_patient})")

            sfdoc = _find_sfdoc(form)
            if sfdoc is None:
                log.error("SFDoc subform not found on this record.")
                file_queue.task_done()
                continue

            # Compute target folder and move file
            nom_dossier_relatif = calculer_dossier_studiovision(code_patient, nom_patient, prenom_patient)
            dest_file = move_file_to_patient(fichier, nom_dossier_relatif)

            if not dest_file:
                file_queue.task_done()
                continue

            chemin_relatif = f"\\{nom_dossier_relatif}\\{dest_file.name}"

            # Safety check: ensure the patient has not changed during the copy
            code_patient_actuel = str(form.Controls("Code patient").Value)
            if code_patient_actuel != code_patient:
                log.error("Patient changed during file copy. Aborting insertion.")
                file_queue.task_done()
                continue

            # Insert record via the GUI (bypasses pyodbc)
            log.info("Inserting record into Studio Vision grid...")
            rs = sfdoc.Recordset
            rs.AddNew()

            rs.Fields("code patient").Value = code_patient
            rs.Fields("Date").Value = datetime.now()
            rs.Fields("DESCRIPTIONS").Value = "Image importée"
            rs.Fields("TEXTE").Value = chemin_relatif
            rs.Fields("Photo externe").Value = chemin_relatif
            rs.Fields("TypeVW").Value = 99

            rs.Update()
            log.info("Record inserted successfully.")

        except Exception as e:
            log.error(f"COM processing error: {e}")

        finally:
            file_queue.task_done()

class ImageProducer(FileSystemEventHandler):
    def __init__(self, file_queue: queue.Queue):
        super().__init__()
        self._queue = file_queue

    def on_created(self, event):
        if event.is_directory:
            return
        fichier = Path(event.src_path)
        if fichier.suffix.lower() in WATCHED_EXTENSIONS:
            self._queue.put(fichier)

def start_watchdog(file_queue: queue.Queue):
    DOSSIER_SOURCE.mkdir(parents=True, exist_ok=True)
    DOSSIER_ORPHELINS.mkdir(parents=True, exist_ok=True)

    observer = Observer()
    observer.schedule(ImageProducer(file_queue), str(DOSSIER_SOURCE), recursive=False)
    observer.start()
    log.info(f"Watching: {DOSSIER_SOURCE}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def main():
    if not WIN32_AVAILABLE:
        return

    print("Router V6 starting...")
    file_queue = queue.Queue()

    # Worker thread handles Access GUI interactions
    t = threading.Thread(target=worker, args=(file_queue,), daemon=True)
    t.start()

    # Main thread watches the local source folder
    start_watchdog(file_queue)

if __name__ == "__main__":
    main()