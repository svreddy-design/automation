"""OpenDental configuration utilities.
Timing, tab order, element locators, and path detection."""

import json
import os
import platform
from core.patient import FIELD_ORDER

DEFAULT_TIMING = {
    "typing_interval_ms": 50,
    "field_delay_ms": 200,
    "action_delay_ms": 500,
    "typing_pause_ms": 30,
    "app_load_delay_s": 12,
    "login_delay_s": 6,
    "batch_patient_delay_s": 3,
}

# Common OpenDental install paths on Windows
OPENDENTAL_PATHS = [
    r"C:\Program Files\Open Dental\OpenDental.exe",
    r"C:\Program Files (x86)\Open Dental\OpenDental.exe",
    r"C:\OpenDental\OpenDental.exe",
]

# Default UI element locators — override in config.json for different OD versions
DEFAULT_LOCATORS = {
    "select_patient_btn": {"title": "Select Patient", "control_type": "SplitButton"},
    "add_pt_btn": {"title_re": ".*Add Pt.*", "control_type": "Button"},
    "save_btn": {"title": "Save", "control_type": "Button"},
    "acknowledge_btn": {"title": "Acknowledge", "control_type": "Button"},
    "ok_btn": {"title": "OK", "control_type": "Button"},
    # Patient form fields
    "last_name_field": {"title": "Last Name", "control_type": "Edit"},
    "first_name_field": {"title": "First Name", "control_type": "Edit"},
    "middle_initial_field": {"title": "Middle Initial", "control_type": "Edit"},
    "birthdate_field": {"title": "Birthdate", "control_type": "Edit"},
    "ssn_field": {"title": "SS#", "control_type": "Edit"},
    "address_field": {"title": "Address", "control_type": "Edit"},
    "city_field": {"title": "City", "control_type": "Edit"},
    "state_field": {"title": "ST", "control_type": "Edit"},
    "zip_field": {"title": "Zip", "control_type": "Edit"},
    "phone_field": {"title": "Home Phone", "control_type": "Edit"},
}


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


def load_locators(config_path="config.json"):
    """Load UI element locators from config. Allows per-OD-version overrides.
    Practices with different OD versions just update config.json, not code."""
    locators = dict(DEFAULT_LOCATORS)
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
                custom = data.get("locators", None)
                if custom and isinstance(custom, dict):
                    locators.update(custom)
        except (json.JSONDecodeError, IOError):
            pass
    return locators
