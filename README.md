# Studiovision-Autosync

Automatic image routing script for [StudioVision](https://www.studiodentaire.com/) — a practice management software for ophthalmologists.  
When a medical imaging device saves a photo, the script detects it, identifies the open patient in StudioVision, moves the file to the correct patient folder on the network drive, and inserts a record in the Access database so the image appears immediately in the patient's file.

---

## Scripts

Several variants are provided in `src/`. They share the same core logic and configuration constants.

| File | Location | Description |
|---|---|---|
| `studiovision_monitor.py` | `src/` | Base version. Watches a flat source folder for new images. |
| `windows7.py` | `src/` | Same as above, using `typing.Optional` for Python 3.9 / Windows 7 compatibility. |
| `box2.py` | `src/` | Extended version with **Nidek device support**. |
| `studiovision_monitorV2.py` | `src/Version 2/` | Improved base version with **batched UI refresh** and **SFDoc-only requery**. |
| `windows7V2.py` | `src/Version 2/` | V2 improvements ported to Python 3.9 / Windows 7 compatible syntax. |
| `box2V2.py` | `src/Version 2/` | Nidek support + V2 improvements (batched refresh, SFDoc-only requery). |
| `studiovision_monitorV3.py` | `src/Version 3/` | Base version with all V3 improvements. |
| `windows7V3.py` | `src/Version 3/` | Python 3.9 / Windows 7 compatible, V3 improvements. |
| `box1V3.py` | `src/Version 3/` | Standard device (Box 1) with all V3 improvements. |
| `box2V3.py` | `src/Version 3/` | Nidek OCT device (Box 2) + all V3 improvements. |
| `studiovision_monitorV4.py` | `src/Version 4/` | **Latest.** Base version with all V4 improvements (system tray, notifications). |
| `windows7V4.py` | `src/Version 4/` | **Latest.** Python 3.9 / Windows 7 compatible, V4 improvements. |
| `box1V4.py` | `src/Version 4/` | **Latest.** Standard device (Box 1) with all V4 improvements. |
| `box2V4.py` | `src/Version 4/` | **Latest.** Nidek OCT device (Box 2) + all V4 improvements. |

---

## How it works

1. **Watchdog** monitors `SOURCE_DIR` for new image and document files (recursively).
2. Each detected file is pushed to a queue and picked up by a background worker thread.
3. The worker waits until the device has finished writing the file (lock-check with retries).
4. It polls the active StudioVision Access form via COM to get the current patient (code, last name, first name).
5. It queries `PUBLIC.MDB` to resolve the patient's folder on the network drive using an existing `Photo externe` entry.
6. It moves the image into that folder, appending a timestamp suffix on name conflict.
7. It inserts a new row into `PUBLIC.MDB` so StudioVision registers the image.
8. It requeues the `SFDoc` subform (with a `Refresh()` fallback) and moves to the last record so the new image is immediately visible.
9. If no patient is found within the configured timeout, the file is moved to the orphan folder.

---

## Nidek device support (`box2.py`, `box2V2.py`, `box2V3.py`, `box2V4.py`)

Nidek devices save scans as a set of files inside a sub-folder (`SOURCE_DIR/<device>/<scan>/`). The box2 variants handle this layout:

- Waits 2 seconds after the first file event to let the full scan land.
- Deletes XML sidecar files automatically.
- Keeps only the **largest image** in the scan folder; all others (thumbnails) are deleted.
- Cleans up the scan folder and its parent once the main image has been processed.
- Tracks already-processed scan folders to drop any residual files that arrive late in the queue.

Files not inside a Nidek sub-folder are processed normally (same as the base version).

---

## Version 4 improvements (`src/Version 4/`)

All four V4 scripts (`studiovision_monitorV4.py`, `windows7V4.py`, `box1V4.py`, `box2V4.py`) are the latest iteration and include all V3 improvements plus the following:

### System tray icon
A persistent icon appears in the Windows notification area (pystray + Pillow):
- **Blue** when idle and ready.
- **Green** while a file transfer is in progress.
- Right-click menu provides a read-only status label, an "Open logs" action, and a "Quit" action.

### Windows toast notifications
Non-blocking toast notifications (~3 seconds) are shown at key events:
- When a transfer starts (first file of a burst).
- When a burst completes, with a count of files processed.
- On errors (file locked, DB insert failure, orphan file).

---

## Version 3 improvements (`src/Version 3/`)

All four V3 scripts (`studiovision_monitorV3.py`, `windows7V3.py`, `box1V3.py`, `box2V3.py`) are the latest iteration and include all previous improvements plus the following:

### Network share wait
At startup, the script blocks until `SOURCE_DIR` is accessible. For UNC/network paths (`\\server\share`), it retries every 10 seconds and logs a warning on each failed attempt.

### Auto-reconnect observer loop
The main loop monitors the Watchdog observer. If it dies (e.g. due to a temporary network drop), the script automatically stops the old observer, waits for the share to come back, and restarts a fresh observer.

### Source directory cleanup at startup
`clear_source_dir()` is called once after the network share is confirmed reachable. It deletes all files and sub-folders left over in `SOURCE_DIR` from a previous session.

### Sleep prevention
`prevent_sleep()` calls `SetThreadExecutionState` to prevent Windows from sleeping or turning off the display while the script is running.

### Burst debounce with patient-code guard
UI refresh is deferred until the queue has been idle for **1.5 seconds**, reducing the number of COM calls during rapid multi-file acquisitions. The patient code captured at insert time is compared against the active patient at refresh time — if the operator navigated away during the burst, the refresh is skipped.

### Dirty-state guard
Before calling `Requery()` on the SFDoc subform, the script checks `form.Dirty`. If the parent form is in edit mode, it clears `Dirty` first to prevent Access from raising a save-prompt dialog.

### Requery retry loop
`Requery()` on `SFDoc` is retried up to 3 times (0.5 s between attempts) before falling back to `Refresh()`.

### Centralized log file
Logs are written to `~/studiovision/image_router.log` (the `studiovision` folder in the user's home directory).

### Document file support
The watched extensions and `EXAM_DESCRIPTION` mapping include document formats:

| Extension | Description inserted |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| `.pdf`, `.rtf`, `.doc`, `.docx`, `.odt` | `Document` |
| all others | `Image` |

---

## Requirements

- **Windows only** — requires `win32com` (COM automation) and `pyodbc` (Access ODBC driver).
- Python 3.10+ (all scripts except `windows7*.py`) or Python 3.9+ (`windows7.py`, `windows7V2.py`, `windows7V3.py`, `windows7V4.py`).
- Microsoft Access ODBC driver installed on the machine.
- `pystray` and `Pillow` for the system tray icon (V4 only — the script falls back to headless mode if unavailable).

| Package | Purpose | Versions |
|---|---|---|
| `watchdog` | File system monitoring | All |
| `pyodbc` | Access database connection via ODBC | All |
| `pywin32` | COM automation for interacting with Access (`win32com`, `pythoncom`) | All |
| `pystray` | System tray icon | V4 only |
| `Pillow` | Icon image generation | V4 only |

### Installation — PowerShell (recommended)

> **Run PowerShell as Administrator** the first time to allow script execution if needed.

**Install all dependencies at once (V4 — full install):**

```powershell
pip install -r requirements.txt
```

**Or install packages individually:**

```powershell
# Core packages — required for all versions
pip install watchdog
pip install pyodbc
pip install pywin32

# Post-install step required for pywin32 (run once after install)
python -m pywin32_postinstall -install

# Optional — only needed for Version 4 system tray / notifications
pip install pystray
pip install Pillow
```

**Verify the installation:**

```powershell
python -c "import watchdog, pyodbc, pythoncom, win32com; print('Core OK')"
python -c "import pystray, PIL; print('Tray OK')"
```

**Upgrade all packages to the latest version:**

```powershell
pip install --upgrade watchdog pyodbc pywin32 pystray Pillow
```

---

## Configuration

Set the following paths at the top of whichever script you run:

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Folder watched for new images (shared by the imaging device) |
| `ORPHAN_DIR` | Destination for files that could not be matched to a patient |
| `DEST_PHOTOS` | Root of the patient photo folders on the network drive |
| `PUBLIC_MDB` | Path to `PUBLIC.MDB` (StudioVision shared database) |
| `DOCUM_MDB` | Path to `DOCUM.MDB` (reserved, currently unused) |

Other tunable constants:

| Constant | Default | Description |
|---|---|---|
| `FILE_LOCK_RETRY_DELAY` | `3` s | Delay between retries when a file is still locked |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Max retries before giving up on a locked file |
| `PATIENT_POLL_INTERVAL` | `3` s | How often to poll Access for an open patient |
| `PATIENT_WAIT_TIMEOUT` | `900` s | Time before orphaning a file if no patient is found (15 min) |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Name of the Access subform listing documents — update if renamed |

---

## Watched extensions

The following file extensions are monitored by default:

`.jpg`, `.jpeg`, `.jfif`, `.png`, `.bmp`, `.tif`, `.tiff`, `.dcm`, `.pdf`, `.rtf`, `.doc`, `.docx`, `.odt`

To add or remove extensions, edit `WATCHED_EXTENSIONS` and update `EXAM_DESCRIPTION` accordingly.

---

## Running

```bash
# Version 4 — recommended
python "src/Version 4/box2V4.py"                    # Nidek OCT device (Box 2) + all V4 improvements
python "src/Version 4/box1V4.py"                    # Standard device (Box 1) + all V4 improvements
python "src/Version 4/studiovision_monitorV4.py"    # Base version with V4 improvements
python "src/Version 4/windows7V4.py"                # Python 3.9 / Windows 7 compatible, V4 improvements

# Version 3
python "src/Version 3/box2V3.py"
python "src/Version 3/box1V3.py"
python "src/Version 3/studiovision_monitorV3.py"
python "src/Version 3/windows7V3.py"

# Version 2
python "src/Version 2/box2V2.py"
python "src/Version 2/studiovision_monitorV2.py"
python "src/Version 2/windows7V2.py"

# Version 1
python src/box2.py
python src/studiovision_monitor.py
python src/windows7.py
```

Logs are written to both the console and `~/studiovision/image_router.log` (V3/V4) or `image_router.log` in the working directory (V1/V2).  
Stop with `Ctrl+C`, or use the **Quit** menu item in the system tray (V4) — the script will finish processing any remaining queued files before exiting.

---

## Patient folder resolution

The script finds the patient folder by querying `PUBLIC.MDB` for an existing `Photo externe` entry for the same patient code. That field stores a relative path:

```
\<group_folder>\<patient_folder>\filename.jpg
```

The group and patient folder names are extracted and combined with `DEST_PHOTOS` to build the absolute path on disk.

---

## Orphan files

A file is moved to `ORPHAN_DIR` when:

- No patient is open in StudioVision within the configured timeout.
- The patient folder cannot be resolved from the database.

All orphan events are logged as warnings and must be handled manually.

---

## Technical notes

- `pythoncom.CoInitialize()` / `CoUninitialize()` are called on the worker thread — COM objects cannot be shared across threads.
- `DOCUM.MDB` is read-only for inserts; all writes go to `PUBLIC.MDB`.
- In V4, pystray requires the main thread on Windows. The observer and worker
