# TIA Auto Save

**TIA Auto Save** is a Python-based GUI tool (with an optional `.exe` version) that automatically saves an active Siemens TIA Portal project at user-defined intervals. It connects to existing TIA Portal processes using the TIA Openness API and helps prevent project loss by regularly triggering the `Save()` method.

## Features

- Auto-save active TIA Portal projects.
- Select running TIA Portal processes and view project paths.
- Define save intervals (in minutes).
- Visual countdown with progress bar and time remaining.
- Start/Stop saving with a single click.
- DLL auto-loader based on selected TIA Portal version (e.g., V18).


## Requirements

- Windows with **TIA Portal V1x** installed (tested with V18).
- .NET Framework installed.
- Python 3.8+  
- Python packages:  
  - `pythonnet`
  - `schedule`

You can install dependencies via:

```bash
pip install -r requirements.txt
```

## How to Run

### Option 1: Python Script

1. Open a command prompt in the script folder.

2. Run:

```bash
python TIA-Auto-Save.py
```

### Option 2: Executable

1. Run TIA-Auto-Save.exe.

## Usage Instructions

1. Enter your installed TIA Portal version (e.g., 18) and click Refresh.

2. Select a running process/project from the dropdown.

3. Set the auto-save interval (in minutes).

4. Click Start Saving to begin periodic saves.

5. You’ll see the progress bar and time left until the next save.

⚠️ If TIA is online, saving may fail due to limitations in the Openness API. The tool will notify you.