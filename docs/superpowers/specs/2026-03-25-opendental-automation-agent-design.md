# OpenDental Automation Agent — Enhanced Desktop Bot

**Date:** 2026-03-25
**Approach:** Enhanced PyWinAuto Bot (Approach A)
**OpenDental Version:** Tested against OpenDental 24.x (tab order may vary — see calibration note)

## Problem

The current bot ([app.py](../../../app.py)) only enters First Name and Last Name into OpenDental. We need full basic demographics, CSV batch import, and a proper project structure with CLAUDE.md and automation skills.

## Design

### 1. Enhanced Patient Data Model

Fields to automate in OpenDental's "Add Patient" form:

| Field | Key | Type | Required | Default | Notes |
|-------|-----|------|----------|---------|-------|
| Last Name | `last_name` | text | yes | "Doe" | Already implemented |
| First Name | `first_name` | text | yes | "John" | Already implemented |
| Middle Initial | `middle_initial` | text | no | "" | Single character, skip if empty |
| Preferred Name | `preferred_name` | text | no | "" | Skip if empty |
| Gender | `gender` | dropdown | no | "" | Values: "Male", "Female", "Unknown". Use arrow keys to select. |
| Date of Birth | `dob` | date | no | "" | Format: MM/DD/YYYY |
| SSN | `ssn` | text | no | "" | 9 digits only, masked in GUI with `show="*"` |
| Address | `address` | text | no | "" | Street address line 1 |
| City | `city` | text | no | "" | City name |
| State | `state` | text | no | "" | 2-letter state code |
| Zip | `zip` | text | no | "" | 5-digit zip |
| Phone (Home) | `phone` | text | no | "" | 10 digits, no formatting |

**Tab order calibration:** The tab order above reflects OpenDental 24.x defaults. A `tab_order` list in `config.json` allows reordering fields if the target version differs. The user can also run a "calibration mode" that just tabs through fields and lets them verify the order.

**Empty/optional fields:** When a field value is empty, the bot sends `Tab` to skip to the next field without typing anything.

### 2. Architecture Changes

```
automation/
├── app.py                    # Enhanced GUI with all demographic fields
├── core/
│   ├── __init__.py
│   ├── opendental.py         # OpenDental-specific automation logic (extracted)
│   ├── patient.py            # Patient dataclass + validation
│   └── csv_import.py         # CSV patient data reader
├── config.json               # Persistent settings + timing + tab order
├── CLAUDE.md                 # Project instructions for Claude
├── requirements.txt          # Updated dependencies (CSV only, no Excel)
└── sample_patients.csv       # Example CSV template
```

**Why `core/` not `automation/`:** The project root is already named `automation/`. Using `core/` avoids the confusing `automation/automation/` nesting and keeps imports clean (`from core.opendental import ...`).

### 3. GUI Layout (Enhanced)

Window resized to 650x850 to fit new fields.

```
+------------------------------------------+
| Practice Management Bot PRO              |
| Automates legacy dental software         |
+------------------------------------------+
| [Patient Info - Scrollable Frame]        |
|  Last Name:    [___________]             |
|  First Name:   [___________]             |
|  Middle Init:  [___]                     |
|  Preferred:    [___________]             |
|  Gender:       [dropdown: M/F/Unknown]   |
|  DOB:          [MM/DD/YYYY]              |
|  SSN:          [*********] (masked)      |
|  Address:      [___________]             |
|  City:         [___________]             |
|  State:        [__]  Zip: [_____]        |
|  Phone:        [___________]             |
+------------------------------------------+
| [Import CSV]  | Status: Ready            |
+------------------------------------------+
| Target App: [path]  [Browse]             |
+------------------------------------------+
| [Run PyWinAuto]    [Run PyAutoGUI]       |
+------------------------------------------+
```

### 4. CSV Import

**Exact CSV headers** (matching the `key` column in Section 1):
```csv
last_name,first_name,middle_initial,preferred_name,gender,dob,ssn,address,city,state,zip,phone
Doe,John,A,,Male,01/15/1990,123456789,123 Main St,Austin,TX,78701,5125551234
Smith,Jane,,,Female,03/22/1985,,456 Oak Ave,Dallas,TX,75201,2145559876
```

**Encoding:** UTF-8.

**Error handling per row:**
- Missing required fields (`last_name`, `first_name`): skip the row, log warning.
- Invalid data (bad date format, SSN not 9 digits, state not 2 chars): skip the invalid field, enter remaining fields, log warning.
- Never abort the batch — always continue to next row.

**Batch progress tracking:**
- Write a `batch_log.csv` after each patient with columns: `row_number, last_name, first_name, status, error`.
- Status values: `success`, `partial` (some fields skipped), `skipped` (missing required).
- On re-run, the bot reads `batch_log.csv` and skips rows already marked `success`, preventing duplicates.

**SSN in CSV:** SSN values are read from disk and typed but never logged to `batch_log.csv` or displayed in the status bar. GUI masks SSN input with `show="*"`.

### 5. OpenDental Field Entry Strategy

Both PyWinAuto and PyAutoGUI engines use the same strategy since OpenDental field entry is keyboard-driven:

1. Navigate to Add Patient form (Ctrl+S → Alt+A, as currently implemented)
2. Cursor lands on Last Name field by default
3. For each field in tab order:
   - If value is non-empty: type it with `pyautogui.write(value, interval=0.05)`
   - If value is empty: do nothing (just tab past)
   - Press `Tab` to advance to next field
4. **Gender dropdown:** Press `Down` arrow key N times (0=Male, 1=Female, 2=Unknown) rather than typing letters
5. **DOB field:** Type digits as `MMDDYYYY` (OpenDental auto-formats with slashes)
6. Press `Enter` to save

**Timing (configurable in config.json):**
```json
{
  "typing_interval_ms": 50,
  "field_delay_ms": 500,
  "app_load_delay_s": 7,
  "login_delay_s": 6,
  "batch_patient_delay_s": 3
}
```

**PyWinAuto difference:** PyWinAuto connects to the process via UIA backend for window focus management. Field entry still uses `pyautogui` keystrokes since OpenDental's WinForms controls respond reliably to keyboard input. The PyWinAuto advantage is window management (connect, focus, re-grab on focus loss).

### 6. Error Handling

- Before each field entry: verify window title contains "Patient" (re-grab if not)
- If window focus lost mid-entry: attempt re-focus up to 3 times, then abort current patient
- Log each step to status bar: "Entering Last Name... Entering DOB..."
- On exception: report which field failed, continue to next patient in batch mode

### 7. CLAUDE.md Content

```markdown
# OpenDental Automation Agent

## Purpose
Desktop GUI automation bot for OpenDental dental practice management software.
Enters patient demographic data via keyboard automation (PyWinAuto + PyAutoGUI).

## Quick Start
pip install -r requirements.txt
python app.py

## Architecture
- app.py — CustomTkinter GUI, launches automation threads
- core/opendental.py — OpenDental navigation + field entry logic
- core/patient.py — Patient dataclass, validation
- core/csv_import.py — CSV reader for batch import

## How Automation Works
1. Launch OpenDental via configured path
2. Bypass Admin login (Enter)
3. Ctrl+S → Select Patient → Alt+A → Add Patient
4. Tab through fields, type values via pyautogui
5. Enter to save

## Conventions
- All timing delays are configurable in config.json
- Tab order is configurable for different OpenDental versions
- Both engines (PyWinAuto/PyAutoGUI) share the same field entry logic
- Windows-only for PyWinAuto; PyAutoGUI works cross-platform

## Build
pyinstaller --noconsole --onefile --windowed --name "PracticeManagementBotPRO" app.py
```

### 8. Skills

**Skill: `opendental-automate`**
- Description: Launch the OpenDental automation GUI
- Action: `python app.py`

**Skill: `opendental-add-patient`**
- Description: Add a single patient from CLI
- Args: `--first "John" --last "Doe" [--dob "01/15/1990"] [--gender "Male"] ...`
- Action: Import `core.opendental` and run headless (no GUI)

## Out of Scope (for now)

- OpenDental REST API integration
- Web/Playwright automation
- Insurance data entry
- Appointment scheduling
- Multi-screen workflows beyond Add Patient
- Excel file support (CSV only)
- SSN encryption at rest

## Success Criteria

1. Bot can enter all 12 demographic fields into OpenDental Add Patient form
2. CSV import processes a batch file, skipping invalid rows, logging results to `batch_log.csv`
3. Both PyWinAuto and PyAutoGUI engines support all fields via shared entry logic
4. CLAUDE.md exists with architecture, conventions, and build instructions
5. Clean module structure: `core/opendental.py`, `core/patient.py`, `core/csv_import.py`
6. Configurable timing and tab order in `config.json`
7. SSN masked in GUI and excluded from logs
