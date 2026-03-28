# OpenDental Patient Entry Bot

Automates patient data entry into OpenDental dental practice management software. Enter patients one at a time through the GUI or batch import from a CSV file.

## What it does

- Opens OpenDental and navigates to the Add Patient form
- Fills in patient details (name, DOB, address, phone, SSN, etc.)
- Searches for duplicates before adding
- Handles popups automatically (works with both Trial and Paid versions)
- Supports CSV batch import with resume capability
- Asks for confirmation before each entry

## Requirements

- Windows PC
- OpenDental installed and running
- Python 3.10 or higher

## Setup

```
git clone https://github.com/svreddy-design/automation.git
cd automation
pip install -r requirements.txt
```

## Usage

### Single Patient

1. Open OpenDental on your machine
2. Run the bot

```
python app.py
```

3. Fill in the patient details in the form
4. Click the green "Enter Patient in OpenDental" button
5. Click "Yes" on the confirmation dialog
6. The bot will open Select Patient, search, click Add Pt, fill the form, and save

### CSV Batch Import

1. Prepare a CSV file with these headers:

```
last_name,first_name,middle_initial,preferred_name,gender,dob,ssn,address,city,state,zip,phone
```

2. See `sample_patients.csv` for an example
3. Click "Import CSV Batch" in the app and select your file
4. The bot enters each patient one by one
5. Progress is saved to `batch_log.csv` so you can resume if interrupted

## Configuration

Create a `config.json` file to customize timing and behavior:

```json
{
  "app_path": "C:\\Program Files (x86)\\Open Dental\\OpenDental.exe",
  "field_delay_ms": 150,
  "typing_interval_ms": 30,
  "app_load_delay_s": 12,
  "batch_patient_delay_s": 3
}
```

## How it works

The bot uses a hybrid approach:
- pywinauto detects windows and finds controls by their internal IDs
- pyautogui handles toolbar clicks and keyboard input where pywinauto cannot

OpenDental uses custom controls that are partially invisible to standard Windows automation tools. The bot works around this by combining both tools for maximum reliability.

## Patient Fields

| Field | Required |
|-------|----------|
| Last Name | Yes |
| First Name | Yes |
| Middle Initial | No |
| Preferred Name | No |
| Date of Birth | No |
| SSN | No |
| Gender | No |
| Address | No |
| City | No |
| State | No |
| Zip | No |
| Phone | No |

## Build Standalone Exe

```
pip install pyinstaller
pyinstaller --noconsole --onefile --windowed --name "PatientEntryBot" app.py
```

The exe will be in the `dist` folder.
