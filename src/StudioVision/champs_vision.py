"""
champs_vision.py

Scans all open Access forms via the Running Object Table (ROT) and dumps
every control (name, type, value), recursing into subforms and tab pages.
Useful for discovering hidden or nested fields in StudioVision.

Usage:
    python champs_vision.py
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

OUTPUT_FILE = Path(os.path.expanduser("~")) / "Desktop" / "diagnostic_lentilles.txt"
lines = []


def p(text=""):
    print(text)
    lines.append(text)


def dump_controls(form, depth=0):
    """Recursively dumps all controls of a form, including subforms."""
    indent = "  " * depth
    try:
        count = form.Controls.Count
    except Exception as e:
        p(f"{indent}[Could not read controls: {e}]")
        return

    for i in range(count):
        try:
            ctrl = form.Controls(i)
        except Exception:
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

        type_name = {
            100: "Label", 101: "Rectangle", 102: "Line", 103: "Image",
            104: "CommandButton", 105: "ToggleButton", 106: "OptionButton",
            107: "CheckBox", 108: "OptionGroup", 109: "BoundObjectFrame",
            110: "TextBox", 111: "ListBox", 112: "SubForm", 113: "ComboBox",
            114: "ObjectFrame", 118: "PageBreak", 119: "CustomControl",
            122: "Attachment", 123: "TabControl", 124: "Page",
        }.get(ctrl_type, f"Type{ctrl_type}")

        marker = " <-- LENTILLE" if "lentil" in name.lower() else ""
        p(f"{indent}[{i}] {type_name:<15} Name={name!r:<40} Value={value}{marker}")

        if ctrl_type == 112:  # Subform — recurse
            try:
                dump_controls(ctrl.Form, depth + 2)
            except Exception as e:
                p(f"{indent}  [Subform error: {e}]")

        if ctrl_type == 123:  # TabControl — controls are in child Pages
            p(f"{indent}  [TabControl: controls are distributed across its Pages]")


def main():
    pythoncom.CoInitialize()

    p("Studio Vision — Deep Access scanner (via ROT)")
    p()

    rot  = pythoncom.GetRunningObjectTable()
    enum = rot.EnumRunning()

    instances_found = 0

    while True:
        try:
            result = enum.Next()
        except Exception:
            break
        if result is None:
            break

        moniker = result[0] if isinstance(result, tuple) else result

        try:
            ctx  = pythoncom.CreateBindCtx(0)
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue

        if not any(name.lower().endswith(ext) for ext in (".mde", ".mdb", ".accdb")):
            continue

        instances_found += 1
        p(f"Access instance found: {name}")

        try:
            obj      = rot.GetObject(moniker)
            dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
            db_obj   = win32com.client.Dispatch(dispatch)

            try:
                app = db_obj.Application
            except Exception:
                app = db_obj

            fc = app.Forms.Count
            p(f"  Open forms: {fc}")

            for j in range(fc):
                form_name = app.Forms(j).Name
                p(f"  Form: {form_name!r}")
                dump_controls(app.Forms(j), depth=2)

        except Exception as e:
            p(f"  [Connection error: {e}]")

    if instances_found == 0:
        p("ERROR: No Access database (.mde, .mdb) found in memory.")
        p("Make sure Studio Vision is open.")

    p()
    p(f"Scan complete. Output saved to: {OUTPUT_FILE}")

    _save_and_exit()


def _save_and_exit():
    text = "\n".join(lines)
    try:
        OUTPUT_FILE.write_text(text, encoding="utf-8")
    except Exception as e:
        print(f"Save error: {e}")

    try:
        os.startfile(str(OUTPUT_FILE))
    except Exception:
        pass

    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()