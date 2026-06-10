"""
Lens.py
Reads and writes fields in the LENTILLES form of StudioVision via COM.
"""

import sys
try:
    import pythoncom
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is not installed.")
    sys.exit(1)

# Field mapping: logical key -> Access control name
LENS_MAPPING = {
    # Right eye (OD)
    "labo_od":      "CmbLaboD",
    "lentille_od":  "CmbLentilleD",
    "rayon_od":     "CmbRayonD",
    "diametre_od":  "CmbDiametreD",
    "sphere_od":    "CmbSphD",
    "cylindre_od":  "CmbCylD",
    "axe_od":       "CmbAxeD",
    "acuite_od":    "CmbAVD",
    "addition_od":  "CmbAddD",
    "parinaud_od":  "CmbParinoD",

    # Left eye (OG)
    "labo_og":      "CmbLaboG",
    "lentille_og":  "CmbLentilleG",
    "rayon_og":     "CmbRayonG",
    "diametre_og":  "CmbDiametreG",
    "sphere_og":    "CmbSphG",
    "cylindre_og":  "CmbCylG",
    "axe_og":       "CmbAxeG",
    "acuite_og":    "CmbAVG",
    "addition_og":  "CmbAddG",
    "parinaud_og":  "CmbParinoG",

    # General
    "entretien":     "ENTRETIEN",
    "commentaires":  "COMMENTAIRES LENTILLES",
}


def get_access_app():
    """Returns the first Access.Application instance found in the ROT, or None."""
    rot  = pythoncom.GetRunningObjectTable()
    enum = rot.EnumRunning()

    while True:
        result = enum.Next()
        if not result:
            break
        moniker = result[0] if isinstance(result, tuple) else result
        try:
            ctx  = pythoncom.CreateBindCtx(0)
            name = moniker.GetDisplayName(ctx, None)
            if any(name.lower().endswith(ext) for ext in (".mde", ".mdb", ".accdb")):
                obj      = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                db_obj   = win32com.client.Dispatch(dispatch)
                try:
                    return db_obj.Application
                except Exception:
                    return db_obj
        except Exception:
            continue
    return None


def get_lentilles_form(app):
    """Returns the LENTILLES form if open, otherwise None."""
    try:
        for i in range(app.Forms.Count):
            if app.Forms(i).Name == "LENTILLES":
                return app.Forms(i)
    except Exception:
        pass
    return None


def lire_lentilles():
    """Reads all mapped fields from the LENTILLES form. Returns a dict or None on error."""
    app = get_access_app()
    if not app:
        print("Error: no Access database found.")
        return None

    form = get_lentilles_form(app)
    if not form:
        print("Error: form 'LENTILLES' is not open in StudioVision.")
        return None

    data = {}
    for key, control_name in LENS_MAPPING.items():
        try:
            value = form.Controls(control_name).Value
            data[key] = str(value) if value is not None else ""
        except Exception:
            data[key] = "[read error]"

    return data


def ecrire_lentilles(new_data):
    """
    Writes values into the LENTILLES form.
    Args:
        new_data: dict mapping logical keys to values, e.g. {"sphere_od": "-2.50"}
    Returns True on success, False if the form could not be reached.
    """
    app = get_access_app()
    if not app:
        return False

    form = get_lentilles_form(app)
    if not form:
        return False

    for key, value in new_data.items():
        if key in LENS_MAPPING:
            control_name = LENS_MAPPING[key]
            try:
                form.Controls(control_name).Value = str(value)
                print(f"  {key} = {value}")
            except Exception as e:
                print(f"  Write failed for '{key}': {e}")

    return True


if __name__ == "__main__":
    pythoncom.CoInitialize()

    print("--- READ ---")
    lens_data = lire_lentilles()
    if lens_data:
        for key, value in lens_data.items():
            print(f"  {key:<15} : {value}")

    # Uncomment to test writing:
    # print("\n--- WRITE ---")
    # ecrire_lentilles({
    #     "sphere_od": "-1.25",
    #     "sphere_og": "-1.50",
    #     "entretien": "Test",
    # })

    input("\nPress Enter to exit...")