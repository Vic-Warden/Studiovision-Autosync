import sys
import time
from pathlib import Path
from datetime import datetime

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("Error: pywin32 module not found.")
    sys.exit(1)

DOSSIER_SOURCE = Path(r"C:\Users\Box-6\Desktop\Salle 6")

_AC_SUBFORM = 112
SFDOC_SUBFORM_NAME = "SFDoc"

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

def main():
    print("Test: GUI-based record insertion with real source folder")

    # 0. Récupérer une image du dossier source
    if not DOSSIER_SOURCE.exists() or not DOSSIER_SOURCE.is_dir():
        print(f"Error: source folder not found: {DOSSIER_SOURCE}")
        return

    extensions_valides = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
    fichiers = [f for f in DOSSIER_SOURCE.iterdir() if f.is_file() and f.suffix.lower() in extensions_valides]

    if not fichiers:
        print(f"No images found in {DOSSIER_SOURCE}.")
        return

    fichier_a_envoyer = fichiers[0]
    chemin_image = str(fichier_a_envoyer)
    print(f"0. Selected image: {chemin_image}")

    try:
        print("1. Connecting to Studio Vision...")
        access = win32com.client.GetActiveObject("Access.Application")
        form = access.Screen.ActiveForm

        if form is None:
            print("No active form found.")
            return

        code_patient = form.Controls("Code patient").Value
        print(f"   Current patient: {form.Controls('NOM').Value} (Code: {code_patient})")

        print("2. Locating document subform (SFDoc)...")
        sfdoc = _find_sfdoc(form)

        if sfdoc is None:
            print("Subform not found.")
            return

        print("   Subform found.")

        print("3. Adding new record...")
        rs = sfdoc.Recordset
        rs.AddNew()

        rs.Fields("code patient").Value = code_patient
        rs.Fields("Date").Value = datetime.now()
        rs.Fields("DESCRIPTIONS").Value = "IMAGE TEST DEPUIS DOSSIER"
        rs.Fields("TEXTE").Value = chemin_image
        rs.Fields("Photo externe").Value = chemin_image
        rs.Fields("TypeVW").Value = 99

        print("   Saving record...")
        rs.Update()

        print("Success. The test record was added. Check Studio Vision to confirm.")

    except Exception as e:
        print(f"Error: {e}")
        print("This may be caused by a lock or focus issue with the Access interface.")

if __name__ == "__main__":
    main()