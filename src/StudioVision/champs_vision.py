"""
diagnostic_access.py
====================
Run this script WHILE Studio Vision is open and a patient is displayed.
It dumps every control name/value visible in the active form and all
subforms, so we can find exactly where Code patient / NOM / Prénom live.

Usage:
    python diagnostic_access.py
    (or double-click it — output goes to diagnostic_output.txt on the Desktop)
"""

import os
import sys
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    input("ERROR: pywin32 not installed. Press Enter to exit.")
    sys.exit(1)

OUTPUT_FILE = Path(os.path.expanduser("~")) / "Desktop" / "diagnostic_output.txt"
lines = []

def p(text=""):
    print(text)
    lines.append(text)


def dump_controls(form, depth=0):
    indent = "  " * depth
    try:
        count = form.Controls.Count
    except Exception as e:
        p(f"{indent}[Cannot read controls: {e}]")
        return

    for i in range(count):
        try:
            ctrl = form.Controls(i)
        except Exception as e:
            p(f"{indent}[Control {i} error: {e}]")
            continue

        try:
            name = str(ctrl.Name)
        except Exception:
            name = f"<control_{i}>"

        try:
            ctrl_type = int(ctrl.ControlType)
        except Exception:
            ctrl_type = -1

        try:
            value = repr(ctrl.Value)
        except Exception:
            value = "<no value>"

        # ControlType 112 = subform, 109 = text box, 110 = label, etc.
        type_name = {
            100: "Label", 101: "Rectangle", 102: "Line", 103: "Image",
            104: "CommandButton", 105: "ToggleButton", 106: "OptionButton",
            107: "CheckBox", 108: "OptionGroup", 109: "BoundObjectFrame",
            110: "TextBox", 111: "ListBox", 112: "SubForm", 113: "ComboBox",
            114: "ObjectFrame", 118: "PageBreak", 119: "CustomControl",
            122: "Attachment", 123: "NavigationControl",
        }.get(ctrl_type, f"Type{ctrl_type}")

        p(f"{indent}[{i}] {type_name:<15} Name={name!r:<40} Value={value}")

        # Recurse into subforms
        if ctrl_type == 112:
            p(f"{indent}  >>> Entering subform: {name!r}")
            try:
                dump_controls(ctrl.Form, depth + 2)
            except Exception as e:
                p(f"{indent}  [Subform error: {e}]")
            p(f"{indent}  <<< End subform: {name!r}")


def main():
    pythoncom.CoInitialize()

    p("=" * 70)
    p("Studio Vision — Access COM Diagnostic")
    p("=" * 70)
    p()

    try:
        access = win32com.client.GetActiveObject("Access.Application")
        p(f"Access version : {access.Version}")
        p(f"Access visible : {access.Visible}")
        p()
    except Exception as e:
        p(f"ERROR: Cannot connect to Access.Application: {e}")
        p()
        p("Make sure Studio Vision is open and a patient record is displayed.")
        _save_and_exit()
        return

    try:
        form = access.Screen.ActiveForm
        if form is None:
            p("ERROR: access.Screen.ActiveForm is None — no form is active.")
            p("Click on the patient form in Studio Vision and run again.")
            _save_and_exit()
            return
        p(f"Active form name : {form.Name!r}")
        p(f"Active form caption : {getattr(form, 'Caption', '?')!r}")
        p()
    except Exception as e:
        p(f"ERROR reading ActiveForm: {e}")
        _save_and_exit()
        return

    p("--- All controls (recursing into subforms) ---")
    p()
    dump_controls(form)

    p()
    p("=" * 70)
    p("Diagnostic complete.")
    p(f"Output saved to: {OUTPUT_FILE}")

    _save_and_exit()


def _save_and_exit():
    text = "\n".join(lines)
    try:
        OUTPUT_FILE.write_text(text, encoding="utf-8")
        print(f"\nSaved to {OUTPUT_FILE}")
    except Exception as e:
        print(f"Could not save file: {e}")

    try:
        os.startfile(str(OUTPUT_FILE))
    except Exception:
        pass

    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()