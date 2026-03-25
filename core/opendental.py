import pyautogui
import time
import json
import os
from core.patient import Patient, FIELD_ORDER, GENDER_MAP

DEFAULT_TIMING = {
    "typing_interval_ms": 50,
    "field_delay_ms": 500,
    "app_load_delay_s": 7,
    "login_delay_s": 6,
    "batch_patient_delay_s": 3,
}


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


def navigate_to_add_patient(status_callback, timing):
    """Navigate from OpenDental main screen to the Add Patient form.
    status_callback(text, color) updates the GUI status bar."""
    status_callback("Bypassing Open Dental Login...", "yellow")
    pyautogui.press('enter')
    time.sleep(timing["login_delay_s"])

    status_callback("Opening 'Select Patient'...", "yellow")
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2)

    status_callback("Clicking 'Add Pt'...", "yellow")
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
                    # Use arrow keys for dropdown
                    gender_index = GENDER_MAP.get(value.lower(), -1)
                    if gender_index >= 0:
                        for _ in range(gender_index):
                            pyautogui.press('down')
                            time.sleep(0.1)
                elif field_name == "dob":
                    # Type digits only: MMDDYYYY (OpenDental auto-formats)
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


def save_patient(status_callback):
    """Press Enter to save the patient record."""
    status_callback("Saving Patient Profile...", "yellow")
    pyautogui.press('enter')
    time.sleep(2)
