# OpenDental Automation Agent

## Purpose
Desktop GUI automation bot for OpenDental dental practice management software.
Enters patient demographic data via keyboard automation (PyWinAuto + PyAutoGUI).

## Quick Start
```
pip install -r requirements.txt
python app.py
```

## Architecture
- `app.py` — CustomTkinter GUI, launches automation threads
- `core/opendental.py` — OpenDental navigation + field entry logic
- `core/patient.py` — Patient dataclass with validation
- `core/csv_import.py` — CSV reader for batch import with resume support

## How Automation Works
1. Launch OpenDental via configured path
2. Bypass Admin login (Enter key)
3. Ctrl+S to open Select Patient, Alt+A to Add Patient
4. Tab through 12 fields, type values via pyautogui
5. Enter to save

## Patient Fields
last_name*, first_name*, middle_initial, preferred_name, gender, dob, ssn, address, city, state, zip, phone

## CSV Batch Import
- Use `sample_patients.csv` as template
- Headers must match field keys exactly
- Results logged to `batch_log.csv` (resumes on re-run)
- SSN never appears in logs
- OpenDental must already be running for batch mode

## Config (`config.json`)
- `app_path` — path to OpenDental executable
- `typing_interval_ms` — delay between keystrokes (default: 50)
- `field_delay_ms` — delay between fields (default: 500)
- `app_load_delay_s` — wait for app launch (default: 7)
- `login_delay_s` — wait after login bypass (default: 6)
- `batch_patient_delay_s` — delay between patients in batch (default: 3)
- `tab_order` — optional list to reorder fields for different OpenDental versions

## Conventions
- All timing delays are configurable in config.json
- Tab order is configurable for different OpenDental versions
- Both engines (PyWinAuto/PyAutoGUI) share the same field entry logic in `core/opendental.py`
- PyWinAuto handles window management; PyAutoGUI handles keystrokes
- Windows-only for PyWinAuto; PyAutoGUI works cross-platform
- SSN is masked in GUI (`show="*"`) and excluded from all logs

## Build
```
pyinstaller --noconsole --onefile --windowed --name "PracticeManagementBotPRO" app.py
```

## Testing
- Verify on Windows with OpenDental installed
- Use sample_patients.csv for batch testing
- Check batch_log.csv for results
