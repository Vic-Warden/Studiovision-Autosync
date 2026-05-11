# Studiovision-Autosync

Automatic image routing script for [StudioVision](https://www.studiodentaire.com/) — a dental practice management software.  
When a medical imaging device saves a photo, the script detects it, identifies the open patient in StudioVision, moves the file to the correct patient folder on the network drive, and inserts a record in the Access database so the image appears immediately in the patient's file.

---

## Scripts

Several variants are provided in `src/`. They share the same core logic and configuration constants.

| File | Location | Description |
|---|---|---|
| `studiovision_monitor.py` | `src/` | Base version. Watches a flat source folder for new images. |
| `windows7.py` | `src/` | Same as above, using `typing.Optional` for Python 3.9 / Windows 7 compatibility. |
| `box2.py` | `src/` | Extended version with **Nidek device support**. |
| `studiovision_monitorV2.py` | `src/Vesion 2/` | Improved base version with **batched UI refresh** and **SFDoc-only requery**. |
| `windows7V2.py` | `src/Vesion 2/` | V2 improvements ported to Python 3.9 / Windows 7 compatible syntax. |
| `box2V2.py` | `src/Vesion 2/` | Nidek support + V2 improvements (batched refresh, SFDoc-only requery). |
| `studiovision_monitorV3.py` | `src/Version 3/` | **Latest.** Base version with all V3 improvements. |
| `windows7V3.py` | `src/Version 3/` | **Latest.** Python 3.9 / Windows 7 compatible, V3 improvements. |
| `box1V3.py` | `src/Version 3/` | **Latest.** Standard device (Box 1) with all V3 improvements. |
| `box2V3.py` | `src/Version 3/` | **Latest.** Nidek OCT device (Box 2) + all V3 improvements. |

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

## Nidek device support (`box2.py`, `box2V2.py`, `box2V3.py`)

Nidek devices save scans as a set of files inside a sub-folder (`SOURCE_DIR/<device>/<scan>/`). The box2 variants handle this layout:

- Waits 2 seconds after the first file event to let the full scan land.
- Deletes XML sidecar files automatically.
- Keeps only the **largest image** in the scan folder; all others (thumbnails) are deleted.
- Cleans up the scan folder and its parent once the main image has been processed.
- Tracks already-processed scan folders to drop any residual files that arrive late in the queue.

Files not inside a Nidek sub-folder are processed normally (same as the base version).

---

## Version 3 improvements (`src/Version 3/`)

All four V3 scripts (`studiovision_monitorV3.py`, `windows7V3.py`, `box1V3.py`, `box2V3.py`) are the latest iteration and include all previous improvements plus the following:

### Network share wait
At startup, the script blocks until `SOURCE_DIR` is accessible. For UNC/network paths (`\\server\share`), it retries every 10 seconds and logs a warning on each failed attempt, so the program waits silently rather than crashing if the share is temporarily unreachable.

### Auto-reconnect observer loop
The main loop monitors the Watchdog observer. If it dies (e.g. due to a temporary network drop), the script automatically stops the old observer, waits for the share to come back, and restarts a fresh observer — no manual restart required.

### Source directory cleanup at startup
`clear_source_dir()` is called once after the network share is confirmed reachable. It deletes all files and sub-folders left over in `SOURCE_DIR` from a previous session, ensuring a clean slate.

### Sleep prevention
`prevent_sleep()` calls `SetThreadExecutionState` to prevent Windows from sleeping or turning off the display while the script is running.

### Burst debounce with patient-code guard
UI refresh is deferred until the queue has been idle for **1.5 seconds**, reducing the number of COM calls during rapid multi-file acquisitions. The patient code captured at insert time is compared against the active patient at refresh time — if the operator navigated away during the burst, the refresh is skipped to avoid updating the wrong record.

### Dirty-state guard
Before calling `Requery()` on the SFDoc subform, the script checks `form.Dirty`. If the parent form is in edit mode, it clears `Dirty` first to prevent Access from raising a save-prompt dialog.

### Requery retry loop
`Requery()` on `SFDoc` is retried up to 3 times (0.5 s between attempts) before falling back to `Refresh()`.

### Centralized log file
Logs are now written to `~/studiovision/image_router.log` (the `studiovision` folder in the user's home directory) instead of the working directory, making them easier to find on deployment machines.

### Document file support
The watched extensions and `EXAM_DESCRIPTION` mapping now include document formats:

| Extension | Description inserted |
|---|---|
| `.tif`, `.tiff` | `OCT` |
| `.dcm` | `DICOM` |
| `.pdf`, `.rtf`, `.doc`, `.docx`, `.odt` | `Document` |
| all others | `Image` |

---

## Requirements

- **Windows only** — requires `win32com` (COM automation) and `pyodbc` (Access ODBC driver).
- Python 3.10+ (`studiovision_monitor.py`, `box2.py`, `box2V2.py`, `box2V3.py`) or Python 3.9+ (`windows7.py`, `windows7V2.py`).
- Microsoft Access ODBC driver installed on the machine.

```bash
pip install -r requirements.txt
```

| Package | Purpose |
|---|---|
| `watchdog` | File system monitoring |
| `pyodbc` | Access database connection via ODBC |
| `pywin32` | COM automation for interacting with Access |

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

> **Note:** document extensions (`.pdf`, `.rtf`, `.doc`, `.docx`, `.odt`) are only present in `box2V3.py`. Earlier versions watch image extensions only.

To add or remove extensions, edit `WATCHED_EXTENSIONS` and update `EXAM_DESCRIPTION` accordingly.

---

## Running

```bash
# Version 3 — recommended
python "src/Version 3/box2V3.py"           # Nidek OCT device (Box 2) + all V3 improvements
python "src/Version 3/box1V3.py"           # Standard device (Box 1) + all V3 improvements
python "src/Version 3/studiovision_monitorV3.py"  # Base version with V3 improvements
python "src/Version 3/windows7V3.py"       # Python 3.9 / Windows 7 compatible, V3 improvements

# Version 2
python "src/Vesion 2/box2V2.py"
python "src/Vesion 2/studiovision_monitorV2.py"
python "src/Vesion 2/windows7V2.py"        # Windows 7 / Python 3.9

# Version 1
python src/box2.py
python src/studiovision_monitor.py
python src/windows7.py                     # Windows 7 / Python 3.9
```

Logs are written to both the console and `~/studiovision/image_router.log` (V3) or `image_router.log` in the working directory (V1/V2).  
Stop with `Ctrl+C` — the script will finish processing any remaining queued files before exiting.

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
- The `windows7.py` and `windows7V2.py` variants avoid `X | None` union syntax, using `typing.Optional` instead for compatibility with Python 3.9.
- In V3, the Watchdog observer is a `PollingObserver`, which works reliably on network shares (SMB/UNC) where native filesystem events are not propagated to the client.

---

## Deployment

The scripts are packaged as standalone executables using **PyInstaller** and launched automatically at Windows startup via a shortcut in the Startup folder.

### Build the executable

```cmd
cd C:\PATH\TO\src\Version 3
pyinstaller --onefile --noconsole --name PROGRAM_NAME SCRIPT_NAME.py
```

> Replace `SCRIPT_NAME.py` with the desired variant (`box2V3.py`, `box1V3.py`, etc.) and `PROGRAM_NAME` with the chosen executable name.

### Add to Windows Startup (PowerShell)

```powershell
$exe = "C:\PATH\TO\dist\PROGRAM_NAME.exe"
$startup = [System.Environment]::GetFolderPath("Startup")
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut("$startup\PROGRAM_NAME.lnk")
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = "C:\PATH\TO\dist"
$shortcut.Save()
```

### Remove from Startup & stop the process (CMD)

```cmd
taskkill /f /im PROGRAM_NAME.exe /t
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\PROGRAM_NAME.lnk"
```

### Schedule at logon via Task Scheduler (optional alternative to Startup shortcut)

```cmd
schtasks /create /tn "TASK_NAME" /tr "C:\PATH\TO\dist\PROGRAM_NAME.exe" /sc onlogon /delay 0001:30 /rl highest /ru SYSTEM /f
```

### Disable sleep & hibernation (CMD)

```cmd
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
```

### Check Startup folder & scheduled task

```cmd
dir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
schtasks /query /tn "TASK_NAME"
```
