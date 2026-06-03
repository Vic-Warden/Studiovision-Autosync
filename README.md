# Studiovision-Autosync

Automatic image routing script for [StudioVision](https://www.studiodentaire.com/) — a practice management software for ophthalmologists.  
When a medical imaging device saves a file, the script detects it, identifies the open patient in StudioVision, moves the file to the correct patient folder on the network drive, and inserts a record directly into the Access form via COM automation so the image appears immediately in the patient's file.

---

## Final version — `src/StudioVision/`

The three scripts in `src/StudioVision/` are the production-ready, fully deployed version (V6).  
All previous versions (V1–V5) are kept for reference only and should not be used.

| File | Description |
|---|---|
| `OMV6.py` | Router for the OM acquisition box. Inserts records via COM into the OM StudioVision instance. |
| `windows7V6.py` | Router for the HR acquisition box. Compatible with Python 3.8.10 / Windows 7. |
| `nidekV6.py` | Router for the Nidek OCT device. Handles multi-file scan folders. |

### Key differences from earlier versions

- **No database layer** — `pyodbc` and all SQL queries have been removed. Records are inserted directly into the open Access form via `win32com` GUI automation (AddNew → Update → Requery).
- **Multi-instance safe** — when both OM and HR routers run on the same machine, each binds to its own `msaccess.exe` process via the Running Object Table instead of using `GetActiveObject`.
- **Lifecycle management** — the router launches StudioVision itself, monitors its process, and force-kills zombie COM locks on exit.
- **Patient-change guard** — before inserting, verifies the patient open in Access matches the one captured when the file arrived.
- **Periodic catchup scan** — re-scans `SOURCE_DIR` every 2 minutes to recover files missed during downtime.
- **System tray** — persistent icon (blue = idle, green = active) with a right-click menu and toast notifications.

---

## Legacy versions

The following versions are archived and kept for reference only.

| Version | Location | Status |
|---|---|---|
| V1 | `src/` | Legacy — base version, SQL insert via pyodbc. |
| V2 | `src/Version 2/` | Legacy — batched UI refresh, SFDoc-only requery. |
| V3 | `src/Version 3/` | Legacy — auto-reconnect, sleep prevention, dirty-state guard. |
| V4 | `src/Version 4/` | Legacy — system tray, toast notifications. |
| V5 | `src/AutoLaunch/` | Legacy — auto-launch of StudioVision, lifecycle monitoring. |

---

## How it works

1. **Watchdog** (`PollingObserver`) monitors `SOURCE_DIR` for new image and document files.
2. Each detected file is pushed to a queue and consumed by a background worker thread.
3. The worker waits until the file is fully written (lock-check with retries).
4. It polls the active StudioVision Access form via COM to get the current patient (code, last name, first name).
5. It resolves the patient's folder on the network drive using the Studio Vision naming convention (`<code><nom3>.<prenom3>`).
6. It moves the file into that folder, appending a timestamp suffix on name conflict.
7. It inserts a new record into the SFDoc subform via COM (`AddNew → Update → Requery`).
8. If no patient is open within the configured timeout, the file is moved to the orphan folder.

---

## Requirements

- **Windows only** — requires `pywin32` (COM automation).
- Python 3.10+ (`OMV6.py`, `nidekV6.py`) or Python 3.8.10+ (`windows7V6.py`).
- `pystray` and `Pillow` for the system tray icon (falls back to headless mode if unavailable).
- `psutil` for process lifecycle management.

| Package | Purpose |
|---|---|
| `watchdog` | File system monitoring |
| `pywin32` | COM automation (`win32com`, `pythoncom`) |
| `pystray` | System tray icon |
| `Pillow` | Icon image generation |
| `psutil` | Process monitoring |

### Installation

```powershell
pip install -r requirements.txt
```

Or individually:

```powershell
pip install watchdog pywin32 pystray Pillow psutil
python -m pywin32_postinstall -install
```

---

## Configuration

Set the following constants at the top of whichever script you run:

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Folder watched for new files (drop folder of the imaging device) |
| `ORPHAN_DIR` | Destination for files that could not be matched to a patient |
| `DEST_PHOTOS` | Root of the patient photo folders on the network drive |
| `STUDIO_VISION_CMD` | Command used to launch StudioVision (`msaccess.exe /runtime ...`) |

Other tunable constants:

| Constant | Default | Description |
|---|---|---|
| `FILE_LOCK_RETRY_DELAY` | `3` s | Delay between retries when a file is still locked |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Max retries before giving up on a locked file |
| `PATIENT_POLL_INTERVAL` | `3` s | How often to poll Access for an open patient |
| `PATIENT_WAIT_TIMEOUT` | `900` s | Time before orphaning a file if no patient is found (15 min) |
| `CATCHUP_INTERVAL` | `120` s | Interval between periodic source-dir scans |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Name of the Access subform listing documents |

---

## Watched extensions

`.jpg`, `.jpeg`, `.jfif`, `.png`, `.bmp`, `.tif`, `.tiff`, `.dcm`, `.pdf`, `.rtf`, `.doc`, `.docx`, `.odt`

---

## Running

```powershell
# Final version — use these
pythonw "src/StudioVision/OMV6.py"        # OM acquisition box
pythonw "src/StudioVision/windows7V6.py"  # HR acquisition box (Windows 7)
pythonw "src/StudioVision/nidekV6.py"     # Nidek OCT device
```

Logs are written to `~/studiovision/image_router_<name>.log`.  
Stop via the **Quit** item in the system tray, or `Ctrl+C` in headless mode.

---

## Patient folder resolution

The patient folder is derived from the patient identity fields read from the open Access form, using the Studio Vision naming convention:

```
DEST_PHOTOS\<first2digits>.000\<code><nom3>.<prenom3>\
```

Example: patient code `0042`, name `Dupont`, first name `Marie` → `DEST_PHOTOS\00.000\0042dup.mar\`

---

## Orphan files

A file is moved to `ORPHAN_DIR` when:

- No patient is open in StudioVision within the configured timeout.
- The patient folder cannot be resolved on disk.

All orphan events are logged as warnings and must be handled manually.

---

## Technical notes

- `pythoncom.CoInitialize()` / `CoUninitialize()` are called on the worker thread — COM objects cannot be shared across threads.
- When both OM and HR routers run simultaneously on the same machine, each binds to its own `Access.Application` instance by scanning the Running Object Table (ROT) rather than using `GetActiveObject`, which always returns the first registered instance.
- The router launches StudioVision via `subprocess.Popen`, tracks new `msaccess.exe` PIDs, and force-kills them on exit to release COM locks.
- In tray mode, `pystray` requires the main thread; the observer, worker, and lifecycle threads run as daemons.
