# OpenDental Automation Agent

## Purpose
Desktop GUI automation bot for OpenDental dental practice management software.
Enters patient demographic data via hybrid pywinauto + pyautogui automation. **Windows-only.**

## Quick Start
```
pip install -r requirements.txt
python app.py
```

## Architecture
- `app.py` — CustomTkinter GUI, launches automation threads, human-in-the-loop confirmation
- `core/opendental_gui.py` — Hybrid agent: pywinauto for detection + pyautogui for typing/clicks
- `core/opendental.py` — Timing config, element locator config, path detection utilities
- `core/patient.py` — Patient dataclass with validation and PHI masking
- `core/csv_import.py` — CSV reader for batch import with resume support

## 10 Core Principles
1. **Reliability over speed** — wait for elements, never hardcoded sleep
2. **Error recovery** — handle crashes, never leave app in half-posted state
3. **Smart element ID** — auto_id/control_type/title, never screen coordinates
4. **Verify every action** — read back field values, check for error dialogs after save
5. **HIPAA/PHI security** — mask SSN, phone, address, DOB in all logs
6. **Idempotency** — check for duplicate patients before adding
7. **Human-in-the-loop** — confirmation dialog before automation runs
8. **Handle OD versions** — configurable element locators via config.json
9. **Speed control** — configurable realistic delays between actions
10. **State management** — always know which screen we're on before acting

## How Automation Works
1. Connect to running OpenDental (or launch it)
2. Identify current screen via window title analysis (state machine)
3. Auto-dismiss popups/alerts/database selection dialogs
4. Open Select Patient dialog
5. Check for duplicate patient (idempotency)
6. Click "Add Pt" to open Edit Patient form
7. Fill each field via `set_edit_text()` + verify by reading back
8. Click Save button, check for error dialogs

## Patient Fields
last_name*, first_name*, middle_initial, preferred_name, gender, dob, ssn, address, city, state, zip, phone

## PHI Security
- SSN masked as `***` in all logs and GUI
- Phone, address, DOB masked as `[REDACTED]` in logs
- No PHI ever written to batch_log.csv or status console

## CSV Batch Import
- Use `sample_patients.csv` as template
- Headers must match field keys exactly
- Results logged to `batch_log.csv` (resumes on re-run)
- Uses GUI automation for each patient (slower but no DB dependency)
- Human confirmation before batch starts

## Config (`config.json`)
- `app_path` — path to OpenDental executable
- `typing_interval_ms` — delay between keystrokes (default: 50)
- `field_delay_ms` — delay between fields (default: 200)
- `action_delay_ms` — delay between major actions (default: 500)
- `typing_pause_ms` — delay between keystrokes for typed input (default: 30)
- `app_load_delay_s` — wait for app launch (default: 12)
- `login_delay_s` — wait after login bypass (default: 6)
- `batch_patient_delay_s` — delay between patients in batch (default: 3)
- `tab_order` — optional list to reorder fields for different OpenDental versions
- `locators` — optional dict to override UI element locators per OD version

## Element Locators
Default locators work with standard OpenDental. Override in config.json for different versions:
```json
{
  "locators": {
    "save_btn": {"title": "OK", "control_type": "Button"},
    "last_name_field": {"auto_id": "textLName", "control_type": "Edit"}
  }
}
```

## Conventions
- Pure pywinauto — zero pyautogui dependency
- Windows-only (OpenDental is a Windows application)
- All timing delays configurable in config.json
- Element locators configurable for different OD versions
- Every field verified after entry (read-back check)
- SSN masked in GUI (`show="*"`) and all logs

## Build
```
pyinstaller --noconsole --onefile --windowed --name "PracticeManagementBotPRO" app.py
```

## Testing
- Verify on Windows with OpenDental installed
- Use sample_patients.csv for batch testing
- Check batch_log.csv for results
- Use Accessibility Insights or Inspect.exe to verify element locators
