import sys
import logging

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("Error: pywin32 module not found.")
    print("Install it with: pip install pywin32")
    sys.exit(1)

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s"
)
log = logging.getLogger("TestCOM")

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

def lister_tous_les_champs(form) -> None:
    """Lists all controls in the active form."""
    log.info("Step 1: Full form scan")
    log.info(f"Main form: {form.Name}")

    count = 0
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            name = ctrl.Name
            ctrl_type = ctrl.ControlType
            value = "N/A"
            if hasattr(ctrl, 'Value') and ctrl.Value is not None:
                value = str(ctrl.Value)[:50]
            log.info(f"Field: {name:<20} | Type: {ctrl_type:<3} | Value: {value}")
            count += 1
        except Exception:
            pass

    log.info(f"{count} field(s) scanned.")

def recuperer_liste_documents(form) -> None:
    """Extracts the document list from the SFDoc subform."""
    log.info("Step 2: Document/photo lookup")

    sfdoc = _find_sfdoc(form)
    if sfdoc is None:
        log.warning(f"Subform '{SFDOC_SUBFORM_NAME}' not found.")
        return

    log.info(f"Subform '{SFDOC_SUBFORM_NAME}' found.")

    try:
        rs = sfdoc.Recordset.Clone()

        if rs.RecordCount == 0:
            log.info("Subform is empty (no documents for this patient).")
            return

        rs.MoveFirst()
        log.info(f"{rs.RecordCount} document(s) found.")

        index = 1
        while not rs.EOF:
            try:
                photo_ext = rs.Fields("Photo externe").Value
                texte = rs.Fields("TEXTE").Value
                description = rs.Fields("DESCRIPTIONS").Value

                log.info(f"Document {index}:")
                log.info(f"  Description  : {description}")
                log.info(f"  Photo path   : {photo_ext}")
                log.info(f"  Text/Path    : {texte}")

            except Exception as e_field:
                log.warning(f"  Could not read fields for document {index}: {e_field}")
                columns = [f.Name for f in rs.Fields]
                log.info(f"  Available columns: {columns}")

            rs.MoveNext()
            index += 1

    except Exception as e:
        log.error(f"Error reading recordset: {e}")

def main():
    log.info("Connecting to Access...")
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form = access.Screen.ActiveForm

        if form is None:
            log.warning("Studio Vision is open but no form is active.")
            return

        lister_tous_les_champs(form)
        recuperer_liste_documents(form)

    except Exception as e:
        log.error("Could not connect to Studio Vision.")
        log.error("Make sure the application is open on a patient record.")
        log.error(f"Details: {e}")

if __name__ == "__main__":
    main()