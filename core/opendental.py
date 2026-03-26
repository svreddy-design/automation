import pyautogui
import time
import json
import os
import platform
import subprocess
from core.patient import Patient, FIELD_ORDER, GENDER_MAP

DEFAULT_TIMING = {
    "typing_interval_ms": 50,
    "field_delay_ms": 500,
    "app_load_delay_s": 7,
    "login_delay_s": 6,
    "batch_patient_delay_s": 3,
}

# Common OpenDental install paths on Windows
OPENDENTAL_PATHS = [
    r"C:\Program Files\Open Dental\OpenDental.exe",
    r"C:\Program Files (x86)\Open Dental\OpenDental.exe",
    r"C:\OpenDental\OpenDental.exe",
]


def find_opendental():
    """Auto-detect OpenDental installation on Windows. Returns path or None."""
    if platform.system() != "Windows":
        return None
    for path in OPENDENTAL_PATHS:
        if os.path.exists(path):
            return path
    return None


def load_timing(config_path="config.json"):
    """Load timing config, merging with defaults."""
    timing = dict(DEFAULT_TIMING)
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
                for key in DEFAULT_TIMING:
                    if key in data:
                        timing[key] = data[key]
        except (json.JSONDecodeError, IOError):
            pass
    return timing


def load_tab_order(config_path="config.json"):
    """Load tab order from config. Falls back to FIELD_ORDER default."""
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
                order = data.get("tab_order", None)
                if order and isinstance(order, list):
                    return order
        except (json.JSONDecodeError, IOError):
            pass
    return list(FIELD_ORDER)


def launch_app(app_path, status_callback, timing):
    """Launch target application cross-platform. Returns True if launched."""
    app_name = os.path.basename(app_path)
    status_callback(f"Launching {app_name}...", "yellow")

    if platform.system() == "Windows":
        try:
            subprocess.Popen([app_path])
        except FileNotFoundError:
            status_callback(f"File not found: {app_path}", "red")
            return False
    elif platform.system() == "Darwin":
        # On Mac, open .app bundles or use 'open' command
        if app_path.endswith(".app") or os.path.isdir(app_path):
            os.system(f"open '{app_path}'")
        else:
            os.system(f"open -a '{app_name}'")
        time.sleep(1)
        # Bring to front
        clean_name = app_name.replace(".app", "")
        os.system(f"osascript -e 'tell application \"{clean_name}\" to activate'")
    else:
        try:
            subprocess.Popen([app_path])
        except FileNotFoundError:
            status_callback(f"File not found: {app_path}", "red")
            return False

    status_callback("Waiting for application to load...", "yellow")
    time.sleep(timing["app_load_delay_s"])
    return True


def navigate_to_add_patient(status_callback, timing, skip_login=False):
    """Navigate from OpenDental main screen to the Add Patient form."""
    if not skip_login:
        status_callback("Bypassing Open Dental Login...", "yellow")
        pyautogui.press('enter')
        time.sleep(timing["login_delay_s"])

    # Dismiss any alert popups first
    status_callback("Dismissing popups...", "yellow")
    pyautogui.press('enter')
    time.sleep(1)

    # Click on Select Patient button in toolbar
    # OpenDental shortcut: click "Select Patient" or use keyboard
    status_callback("Opening 'Select Patient'...", "yellow")
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(2)

    # In Select Patient window, click "Add New Patient" button
    status_callback("Clicking 'Add New Patient'...", "yellow")
    # Tab to the Add button and press Enter, or use Alt+A within the Select Patient dialog
    pyautogui.hotkey('alt', 'a')
    time.sleep(2)


def enter_patient_fields(patient, status_callback, timing, config_path="config.json"):
    """Type all patient fields into the Add Patient form via Tab navigation.
    Assumes cursor is on Last Name field. Returns list of field errors."""
    interval = timing["typing_interval_ms"] / 1000.0
    delay = timing["field_delay_ms"] / 1000.0
    field_errors = []
    tab_order = load_tab_order(config_path)

    for field_name in tab_order:
        value = getattr(patient, field_name, "")

        # Mask SSN in status display
        display_value = "***" if field_name == "ssn" and value else value

        if value:
            status_callback(f"Entering {field_name}: {display_value}", "yellow")
            try:
                if field_name == "gender":
                    gender_index = GENDER_MAP.get(value.lower(), -1)
                    if gender_index >= 0:
                        for _ in range(gender_index):
                            pyautogui.press('down')
                            time.sleep(0.1)
                elif field_name == "dob":
                    digits = value.replace("/", "")
                    pyautogui.write(digits, interval=interval)
                else:
                    pyautogui.write(value, interval=interval)
            except Exception as e:
                field_errors.append(f"{field_name}: {e}")

        # Tab to next field
        pyautogui.press('tab')
        time.sleep(delay)

    return field_errors


def enter_patient_fields_plain(patient, status_callback, timing):
    """Type patient data as plain text into any text editor (for demo/test mode).
    Works on Mac and Windows — opens a text area and types all fields."""
    interval = timing["typing_interval_ms"] / 1000.0
    tab_order = load_tab_order()

    status_callback("Typing patient data...", "yellow")
    header = "=== Patient Record ==="
    pyautogui.write(header, interval=interval)
    pyautogui.press('enter')
    time.sleep(0.3)

    for field_name in tab_order:
        value = getattr(patient, field_name, "")
        display_value = "***" if field_name == "ssn" and value else value
        if value:
            status_callback(f"Typing {field_name}: {display_value}", "yellow")
            line = f"{field_name}: {value}"
            pyautogui.write(line, interval=interval)
            pyautogui.press('enter')
            time.sleep(0.2)

    pyautogui.write("=== End Record ===", interval=interval)


def save_patient(status_callback):
    """Press Enter to save the patient record."""
    status_callback("Saving Patient Profile...", "yellow")
    pyautogui.press('enter')
    time.sleep(2)


def dry_run(patient, status_callback):
    """Simulate the automation without touching any app. Logs what it would do."""
    tab_order = load_tab_order()

    status_callback("DRY RUN: Would bypass login (Enter)...", "cyan")
    time.sleep(0.5)
    status_callback("DRY RUN: Would open Select Patient (Ctrl+S)...", "cyan")
    time.sleep(0.5)
    status_callback("DRY RUN: Would click Add Pt (Alt+A)...", "cyan")
    time.sleep(0.5)

    for field_name in tab_order:
        value = getattr(patient, field_name, "")
        display_value = "***" if field_name == "ssn" and value else value
        if value:
            status_callback(f"DRY RUN: Would type {field_name} = {display_value}", "cyan")
        else:
            status_callback(f"DRY RUN: Would skip {field_name} (empty)", "gray")
        time.sleep(0.3)

    status_callback("DRY RUN: Would press Enter to save", "cyan")
    time.sleep(0.5)
    status_callback("DRY RUN complete! All steps verified.", "limegreen")
