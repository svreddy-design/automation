# OpenDental Automation Agent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enhance the Practice Management Bot PRO to enter 12 demographic fields into OpenDental, support CSV batch import, and provide clean project structure with CLAUDE.md.

**Architecture:** Extract OpenDental automation logic from app.py into `core/` package. Patient dataclass handles validation. Shared field entry function used by both PyWinAuto and PyAutoGUI engines. GUI expanded with scrollable form for all demographics.

**Tech Stack:** Python 3.8+, CustomTkinter, PyAutoGUI, PyWinAuto (Windows), csv stdlib

**Spec:** `docs/superpowers/specs/2026-03-25-opendental-automation-agent-design.md`

---

## File Structure

| File | Action | Responsibility |
|------|--------|---------------|
| `core/__init__.py` | Create | Package init, exports |
| `core/patient.py` | Create | Patient dataclass + validation |
| `core/csv_import.py` | Create | CSV reader, batch_log writer, resume logic |
| `core/opendental.py` | Create | OpenDental navigation, field entry, window management |
| `app.py` | Modify | Enhanced GUI with 12 fields, CSV import button, use core/ modules |
| `CLAUDE.md` | Create | Project documentation for Claude |
| `sample_patients.csv` | Create | Example CSV template |
| `requirements.txt` | No change | Already has all needed deps (csv is stdlib) |

---

### Task 1: Create Patient Data Model (`core/patient.py`)

**Files:**
- Create: `core/__init__.py`
- Create: `core/patient.py`

- [ ] **Step 1: Create `core/__init__.py`**

```python
# core/__init__.py - empty init
```

- [ ] **Step 2: Create `core/patient.py` with Patient dataclass and validation**

```python
from dataclasses import dataclass, field, fields, asdict
import re

FIELD_ORDER = [
    "last_name", "first_name", "middle_initial", "preferred_name",
    "gender", "dob", "ssn", "address", "city", "state", "zip", "phone"
]

GENDER_MAP = {"male": 0, "female": 1, "unknown": 2}

@dataclass
class Patient:
    last_name: str = ""
    first_name: str = ""
    middle_initial: str = ""
    preferred_name: str = ""
    gender: str = ""
    dob: str = ""
    ssn: str = ""
    address: str = ""
    city: str = ""
    state: str = ""
    zip: str = ""
    phone: str = ""

    def validate(self):
        """Validate fields. Returns (is_valid, errors) tuple.
        is_valid is False only if required fields are missing.
        errors list contains warnings for invalid optional fields."""
        errors = []
        if not self.last_name.strip():
            errors.append("last_name is required")
        if not self.first_name.strip():
            errors.append("first_name is required")
        if self.middle_initial and len(self.middle_initial) > 1:
            errors.append("middle_initial must be single character")
            self.middle_initial = self.middle_initial[0]
        if self.gender and self.gender.lower() not in GENDER_MAP:
            errors.append(f"gender '{self.gender}' invalid, must be Male/Female/Unknown")
            self.gender = ""
        if self.dob and not re.match(r'^\d{2}/\d{2}/\d{4}$', self.dob):
            errors.append(f"dob '{self.dob}' invalid, must be MM/DD/YYYY")
            self.dob = ""
        if self.ssn and not re.match(r'^\d{9}$', self.ssn):
            errors.append(f"ssn invalid, must be 9 digits")
            self.ssn = ""
        if self.state and not re.match(r'^[A-Z]{2}$', self.state.upper()):
            errors.append(f"state '{self.state}' invalid, must be 2-letter code")
            self.state = ""
        if self.zip and not re.match(r'^\d{5}$', self.zip):
            errors.append(f"zip '{self.zip}' invalid, must be 5 digits")
            self.zip = ""
        if self.phone and not re.match(r'^\d{10}$', self.phone):
            errors.append(f"phone '{self.phone}' invalid, must be 10 digits")
            self.phone = ""

        has_required = bool(self.last_name.strip() and self.first_name.strip())
        return has_required, errors

    def to_dict(self):
        return asdict(self)
```

- [ ] **Step 3: Verify file runs without import errors**

Run: `cd /Users/srivardhanreddygutta/automation && python -c "from core.patient import Patient, FIELD_ORDER, GENDER_MAP; p = Patient(last_name='Doe', first_name='John'); print(p.validate())"`
Expected: `(True, [])`

- [ ] **Step 4: Commit**

```bash
git add core/__init__.py core/patient.py
git commit -m "feat: add Patient dataclass with validation in core/patient.py"
```

---

### Task 2: Create CSV Import Module (`core/csv_import.py`)

**Files:**
- Create: `core/csv_import.py`
- Create: `sample_patients.csv`

- [ ] **Step 1: Create `core/csv_import.py`**

```python
import csv
import os
from core.patient import Patient, FIELD_ORDER

def read_patients_csv(filepath):
    """Read patients from CSV file. Returns list of (row_number, Patient, errors) tuples."""
    results = []
    with open(filepath, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=1):
            patient = Patient(
                last_name=row.get('last_name', '').strip(),
                first_name=row.get('first_name', '').strip(),
                middle_initial=row.get('middle_initial', '').strip(),
                preferred_name=row.get('preferred_name', '').strip(),
                gender=row.get('gender', '').strip(),
                dob=row.get('dob', '').strip(),
                ssn=row.get('ssn', '').strip(),
                address=row.get('address', '').strip(),
                city=row.get('city', '').strip(),
                state=row.get('state', '').strip(),
                zip=row.get('zip', '').strip(),
                phone=row.get('phone', '').strip(),
            )
            is_valid, errors = patient.validate()
            results.append((i, patient, is_valid, errors))
    return results


def load_batch_log(log_path):
    """Load previously completed rows from batch_log.csv. Returns set of completed row numbers."""
    completed = set()
    if not os.path.exists(log_path):
        return completed
    with open(log_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get('status') == 'success':
                completed.add(int(row['row_number']))
    return completed


def write_batch_log_entry(log_path, row_number, last_name, first_name, status, error=""):
    """Append a single entry to batch_log.csv. Never logs SSN."""
    file_exists = os.path.exists(log_path)
    with open(log_path, 'a', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['row_number', 'last_name', 'first_name', 'status', 'error'])
        writer.writerow([row_number, last_name, first_name, status, error])
```

- [ ] **Step 2: Create `sample_patients.csv`**

```csv
last_name,first_name,middle_initial,preferred_name,gender,dob,ssn,address,city,state,zip,phone
Doe,John,A,,Male,01/15/1990,123456789,123 Main St,Austin,TX,78701,5125551234
Smith,Jane,,,Female,03/22/1985,,456 Oak Ave,Dallas,TX,75201,2145559876
Johnson,Robert,B,Bob,Male,07/04/1975,987654321,789 Elm Rd,Houston,TX,77001,7135557890
```

- [ ] **Step 3: Verify CSV import works**

Run: `cd /Users/srivardhanreddygutta/automation && python -c "from core.csv_import import read_patients_csv; rows = read_patients_csv('sample_patients.csv'); print(f'{len(rows)} patients loaded'); print(rows[0])"`
Expected: `3 patients loaded` and first patient tuple

- [ ] **Step 4: Commit**

```bash
git add core/csv_import.py sample_patients.csv
git commit -m "feat: add CSV import module with batch logging and sample data"
```

---

### Task 3: Create OpenDental Automation Engine (`core/opendental.py`)

**Files:**
- Create: `core/opendental.py`

- [ ] **Step 1: Create `core/opendental.py` with shared field entry logic**

```python
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
```

- [ ] **Step 2: Verify import works**

Run: `cd /Users/srivardhanreddygutta/automation && python -c "from core.opendental import load_timing, load_tab_order, DEFAULT_TIMING; print(load_timing()); print(load_tab_order())"`
Expected: prints timing dict and tab order list

- [ ] **Step 3: Commit**

```bash
git add core/opendental.py
git commit -m "feat: add OpenDental automation engine with shared field entry logic"
```

---

### Task 4: Enhance GUI with All Demographic Fields (`app.py`)

**Files:**
- Modify: `app.py` (full rewrite of GUI section and automation methods)

- [ ] **Step 1: Rewrite `app.py` with expanded form, CSV import, and core/ integration**

Key changes to make to `app.py`:
1. Add imports: `from core.patient import Patient, FIELD_ORDER` and `from core.opendental import ...` and `from core.csv_import import ...`
2. Resize window to `650x900`
3. Replace the 2-field input_frame with a scrollable frame containing all 12 fields
4. Add gender dropdown using `CTkOptionMenu`
5. Add SSN entry with `show="*"`
6. Add State/Zip on same row
7. Add CSV Import button
8. Replace inline OpenDental logic in `run_pywinauto()` and `run_pyautogui()` with calls to `core.opendental` functions
9. Add `get_patient_from_gui()` method to collect all fields into a Patient object
10. Add `run_csv_batch()` method for batch processing
11. Enhanced `load_config()` and `save_config()` to handle timing settings

The full replacement `app.py`:

```python
import customtkinter as ctk
import pyautogui
import time
import threading
import os
import platform
import subprocess
import json
from datetime import datetime
from tkinter import filedialog

from core.patient import Patient, FIELD_ORDER
from core.opendental import (
    load_timing, load_tab_order, navigate_to_add_patient,
    enter_patient_fields, save_patient
)
from core.csv_import import read_patients_csv, load_batch_log, write_batch_log_entry

# Securely import pywinauto ONLY if the OS is Windows
try:
    if platform.system() == "Windows":
        from pywinauto import Application
except ImportError:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")


class LegacyAutomationBot(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Practice Management Bot PRO")
        self.geometry("650x900")
        self.resizable(False, False)

        self.config_file = "config.json"
        self.app_path = self.load_config()
        self.timing = load_timing(self.config_file)

        # --- Header ---
        self.header_label = ctk.CTkLabel(
            self, text="Practice App Automation",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.header_label.pack(pady=(15, 2))

        self.instructions = ctk.CTkLabel(
            self,
            text="Automates legacy medical/dental management software.",
            text_color="gray"
        )
        self.instructions.pack(pady=(0, 10))

        # --- Scrollable Patient Form ---
        self.form_frame = ctk.CTkScrollableFrame(self, height=350)
        self.form_frame.pack(pady=5, padx=20, fill="x")

        self.entries = {}
        row = 0

        # Last Name
        ctk.CTkLabel(self.form_frame, text="Last Name *").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["last_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["last_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        self.entries["last_name"].insert(0, "Doe")
        row += 1

        # First Name
        ctk.CTkLabel(self.form_frame, text="First Name *").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["first_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["first_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        self.entries["first_name"].insert(0, "John")
        row += 1

        # Middle Initial
        ctk.CTkLabel(self.form_frame, text="Middle Initial").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["middle_initial"] = ctk.CTkEntry(self.form_frame, width=50)
        self.entries["middle_initial"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # Preferred Name
        ctk.CTkLabel(self.form_frame, text="Preferred Name").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["preferred_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["preferred_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        row += 1

        # Gender (dropdown)
        ctk.CTkLabel(self.form_frame, text="Gender").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.gender_var = ctk.StringVar(value="")
        self.entries["gender"] = ctk.CTkOptionMenu(
            self.form_frame, values=["", "Male", "Female", "Unknown"],
            variable=self.gender_var, width=150
        )
        self.entries["gender"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # DOB
        ctk.CTkLabel(self.form_frame, text="Date of Birth").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["dob"] = ctk.CTkEntry(self.form_frame, width=150, placeholder_text="MM/DD/YYYY")
        self.entries["dob"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # SSN (masked)
        ctk.CTkLabel(self.form_frame, text="SSN").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["ssn"] = ctk.CTkEntry(self.form_frame, width=150, show="*")
        self.entries["ssn"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # Address
        ctk.CTkLabel(self.form_frame, text="Address").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["address"] = ctk.CTkEntry(self.form_frame, width=300)
        self.entries["address"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        row += 1

        # City
        ctk.CTkLabel(self.form_frame, text="City").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["city"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["city"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # State + Zip on same row
        ctk.CTkLabel(self.form_frame, text="State").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["state"] = ctk.CTkEntry(self.form_frame, width=50)
        self.entries["state"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        ctk.CTkLabel(self.form_frame, text="Zip").grid(row=row, column=2, padx=8, pady=5, sticky="w")
        self.entries["zip"] = ctk.CTkEntry(self.form_frame, width=80)
        self.entries["zip"].grid(row=row, column=3, padx=8, pady=5, sticky="w")
        row += 1

        # Phone
        ctk.CTkLabel(self.form_frame, text="Phone").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["phone"] = ctk.CTkEntry(self.form_frame, width=150, placeholder_text="10 digits")
        self.entries["phone"].grid(row=row, column=1, padx=8, pady=5, sticky="w")

        # --- CSV Import + Status ---
        self.mid_frame = ctk.CTkFrame(self)
        self.mid_frame.pack(pady=5, padx=20, fill="x")

        self.csv_btn = ctk.CTkButton(
            self.mid_frame, text="Import CSV",
            command=self.import_csv, fg_color="#444", hover_color="#555", width=120
        )
        self.csv_btn.pack(side="left", padx=10, pady=10)

        self.status_label = ctk.CTkLabel(
            self.mid_frame, text="Status: Ready",
            font=ctk.CTkFont(size=13), text_color="limegreen"
        )
        self.status_label.pack(side="left", padx=10, pady=10, expand=True)

        # --- Settings ---
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.pack(pady=5, padx=20, fill="x")

        self.path_label = ctk.CTkLabel(
            self.settings_frame, text="Target Application Path:",
            font=ctk.CTkFont(weight="bold")
        )
        self.path_label.pack(pady=(8, 0), padx=10)

        self.path_display = ctk.CTkEntry(self.settings_frame, width=400)
        self.path_display.pack(pady=4, padx=10)
        self.path_display.insert(0, self.app_path)
        self.path_display.configure(state="disabled")

        self.browse_btn = ctk.CTkButton(
            self.settings_frame, text="Select App Executable",
            command=self.browse_app_path, fg_color="#444", hover_color="#555"
        )
        self.browse_btn.pack(pady=(0, 8))

        # --- Action Buttons ---
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(5, 15), padx=20, fill="x")

        self.pywin_btn = ctk.CTkButton(
            self.button_frame,
            text="Run PyWinAuto\n(Native Elements)",
            font=ctk.CTkFont(size=14, weight="bold"), height=50,
            command=self.start_pywinauto_thread,
            fg_color="#005b96", hover_color="#03396c"
        )
        self.pywin_btn.pack(side="left", expand=True, fill="x", padx=5)

        self.pyauto_btn = ctk.CTkButton(
            self.button_frame,
            text="Run PyAutoGUI\n(Screen/Images)",
            font=ctk.CTkFont(size=14, weight="bold"), height=50,
            command=self.start_pyautogui_thread,
            fg_color="#b33939", hover_color="#cd6133"
        )
        self.pyauto_btn.pack(side="right", expand=True, fill="x", padx=5)

    # ---------- Helpers ----------
    def update_status(self, text, color="limegreen"):
        self.after(0, lambda: self.status_label.configure(
            text=f"Status: {text}", text_color=color))

    def disable_buttons(self):
        for btn in [self.pywin_btn, self.pyauto_btn, self.browse_btn, self.csv_btn]:
            self.after(0, lambda b=btn: b.configure(state="disabled"))

    def enable_buttons(self):
        for btn in [self.pywin_btn, self.pyauto_btn, self.browse_btn, self.csv_btn]:
            self.after(0, lambda b=btn: b.configure(state="normal"))

    def get_patient_from_gui(self):
        """Collect all form fields into a Patient object."""
        vals = {}
        for key, widget in self.entries.items():
            if key == "gender":
                vals[key] = self.gender_var.get()
            else:
                vals[key] = widget.get()
        return Patient(**vals)

    # ---------- Config ----------
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
                    return data.get("app_path", "winword")
            except (json.JSONDecodeError, IOError):
                return "winword"
        return "winword"

    def save_config(self):
        # Preserve existing config keys, update app_path
        data = {}
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        data["app_path"] = self.app_path
        with open(self.config_file, 'w') as f:
            json.dump(data, f, indent=2)

    def browse_app_path(self):
        file_path = filedialog.askopenfilename(
            title="Select Application Executable",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if file_path:
            self.app_path = file_path
            self.path_display.configure(state="normal")
            self.path_display.delete(0, "end")
            self.path_display.insert(0, self.app_path)
            self.path_display.configure(state="disabled")
            self.save_config()
            self.update_status(f"App Path Updated: {os.path.basename(file_path)}", "cyan")

    # ---------- CSV Import ----------
    def import_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Patient CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.disable_buttons()
            threading.Thread(
                target=self.run_csv_batch, args=(file_path,), daemon=True
            ).start()

    def run_csv_batch(self, csv_path):
        """Process CSV file in batch mode.
        NOTE: OpenDental must already be running and focused.
        Batch mode does not launch the app — it enters patients directly."""
        try:
            rows = read_patients_csv(csv_path)
            log_path = os.path.join(os.path.dirname(csv_path), "batch_log.csv")
            completed = load_batch_log(log_path)
            total = len(rows)

            for row_num, patient, is_valid, errors in rows:
                if row_num in completed:
                    self.update_status(f"Skipping row {row_num} (already done)", "cyan")
                    continue
                if not is_valid:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "skipped", "; ".join(errors)
                    )
                    self.update_status(f"Row {row_num}: skipped (missing required)", "orange")
                    continue

                self.update_status(
                    f"Patient {row_num}/{total}: {patient.first_name} {patient.last_name}",
                    "yellow"
                )
                try:
                    self._run_opendental_entry(patient)
                    status = "partial" if errors else "success"
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, status,
                        "; ".join(errors) if errors else ""
                    )
                except Exception as e:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "error", str(e)
                    )
                    self.update_status(f"Row {row_num} error: {e}", "red")

                time.sleep(self.timing["batch_patient_delay_s"])

            self.update_status(f"Batch complete! {total} patients processed.", "limegreen")
        except Exception as e:
            self.update_status(f"CSV Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- Shared OpenDental Entry ----------
    def _run_opendental_entry(self, patient):
        """Enter a single patient into OpenDental (assumes app is focused on main screen)."""
        navigate_to_add_patient(self.update_status, self.timing)
        field_errors = enter_patient_fields(patient, self.update_status, self.timing)
        save_patient(self.update_status)
        if field_errors:
            self.update_status(
                f"Saved with warnings: {', '.join(field_errors)}", "orange"
            )
        else:
            name = f"{patient.first_name} {patient.last_name}"
            self.update_status(f"Patient Saved: {name}", "limegreen")

    # ---------- PyWinAuto Engine ----------
    def start_pywinauto_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pywinauto, daemon=True).start()

    def run_pywinauto(self):
        if platform.system() != "Windows":
            self.update_status("Error: PyWinAuto ONLY works on Windows!", "red")
            self.enable_buttons()
            return

        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                self.enable_buttons()
                return

            app_name = os.path.basename(self.app_path)
            self.update_status(f"Launching {app_name}...", "yellow")

            if self.app_path == "winword":
                subprocess.Popen(["cmd", "/c", "start winword"], shell=True)
            else:
                subprocess.Popen([self.app_path])

            self.update_status("Waiting for application to load...", "yellow")
            time.sleep(self.timing["app_load_delay_s"])

            # Connect via PyWinAuto UIA
            try:
                path = "winword.exe" if self.app_path == "winword" else self.app_path
                app = Application(backend="uia").connect(path=path, timeout=10)
            except Exception:
                app = Application(backend="uia").connect(
                    title_re=f".*{app_name.split('.')[0]}.*", timeout=10
                )

            main_window = app.top_window()
            main_window.set_focus()

            if "opendental" in app_name.lower():
                self._run_opendental_entry(patient)
            else:
                # Fallback: Word document automation (existing behavior)
                self._run_word_automation(app, patient)

        except Exception as e:
            self.update_status(f"Execution Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- PyAutoGUI Engine ----------
    def start_pyautogui_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pyautogui, daemon=True).start()

    def run_pyautogui(self):
        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                self.enable_buttons()
                return

            app_name = os.path.basename(self.app_path)
            self.update_status(f"Launching {app_name}...", "yellow")

            if platform.system() == "Windows":
                pyautogui.hotkey('win', 'r')
                time.sleep(1)
                target = self.app_path if self.app_path != "winword" else "winword"
                pyautogui.typewrite(target)
                pyautogui.press('enter')
            elif platform.system() == "Darwin":
                target_name = app_name if self.app_path != "winword" else "Microsoft Word"
                os.system(f"open -a '{target_name}'")
                time.sleep(1)
                os.system(f"osascript -e 'tell application \"{target_name}\" to activate'")
                time.sleep(2)

            self.update_status(f"Waiting for {app_name} to load...", "yellow")
            time.sleep(self.timing["app_load_delay_s"])

            if "opendental" in app_name.lower():
                self._run_opendental_entry(patient)
            else:
                self.update_status("Non-OpenDental PyAutoGUI mode", "orange")
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.write(
                    f"Patient: {patient.first_name} {patient.last_name}",
                    interval=0.05
                )
                self.update_status("Screen Automation Complete!", "limegreen")

        except Exception as e:
            self.update_status(f"PyAutoGUI Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- Word Fallback (preserved from original) ----------
    def _run_word_automation(self, app, patient):
        self.update_status("Selecting 'Blank Document'...", "yellow")
        pyautogui.press('enter')
        time.sleep(4)

        doc_window = app.top_window()
        doc_window.set_focus()
        time.sleep(1)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        text = (
            f"Automated Patient Entry (Pro Mode)\n"
            f"First Name: {patient.first_name}\n"
            f"Last Name: {patient.last_name}\n"
            f"Entry Time: {timestamp}"
        )
        pyautogui.write(text, interval=0.05)
        time.sleep(1)

        pyautogui.press('f12')
        time.sleep(2)
        safe_fn = "".join(x for x in patient.first_name if x.isalnum())
        safe_ln = "".join(x for x in patient.last_name if x.isalnum())
        pyautogui.write(f"Patient_{safe_fn}_{safe_ln}_{timestamp}")
        time.sleep(1)
        pyautogui.press('enter')
        self.update_status(f"Document saved!", "limegreen")


if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    app = LegacyAutomationBot()
    app.mainloop()
```

- [ ] **Step 2: Verify app.py imports and initializes without error**

Run: `cd /Users/srivardhanreddygutta/automation && python -c "import app; print('app.py imports OK')"`
Expected: `app.py imports OK` (may fail on macOS if tkinter display isn't available — that's fine, we verify the import chain)

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: enhanced GUI with 12 demographic fields, CSV import, and core/ integration"
```

---

### Task 5: Create CLAUDE.md

**Files:**
- Create: `CLAUDE.md`

- [ ] **Step 1: Create `CLAUDE.md`**

```markdown
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

## Config (`config.json`)
- `app_path` — path to OpenDental executable
- `typing_interval_ms` — delay between keystrokes (default: 50)
- `field_delay_ms` — delay between fields (default: 500)
- `app_load_delay_s` — wait for app launch (default: 7)
- `login_delay_s` — wait after login bypass (default: 6)
- `batch_patient_delay_s` — delay between patients in batch (default: 3)

## Conventions
- All timing delays are configurable in config.json
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
```

- [ ] **Step 2: Verify CLAUDE.md exists and has correct header**

Run: `cd /Users/srivardhanreddygutta/automation && head -3 CLAUDE.md`
Expected: `# OpenDental Automation Agent`

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: add CLAUDE.md with project architecture and conventions"
```

---

### Task 6: Create Sample CSV and Verify Full Pipeline

**Files:**
- Verify: `sample_patients.csv` (created in Task 2)
- Verify: all imports work end-to-end

- [ ] **Step 1: Verify full import chain**

Run: `cd /Users/srivardhanreddygutta/automation && python -c "
from core.patient import Patient, FIELD_ORDER, GENDER_MAP
from core.csv_import import read_patients_csv, load_batch_log, write_batch_log_entry
from core.opendental import load_timing, navigate_to_add_patient, enter_patient_fields, save_patient

# Test patient creation + validation
p = Patient(last_name='Test', first_name='User', dob='01/01/2000', gender='Male', ssn='123456789')
valid, errors = p.validate()
print(f'Patient valid: {valid}, errors: {errors}')

# Test CSV read
rows = read_patients_csv('sample_patients.csv')
print(f'CSV rows: {len(rows)}')

# Test timing
t = load_timing()
print(f'Timing keys: {list(t.keys())}')

print('All modules OK!')
"`

Expected: All prints succeed, `All modules OK!`

- [ ] **Step 2: Commit any fixes if needed**

---

### Task 7: Save Memory and Final Cleanup

- [ ] **Step 1: Save project memory**

Write a memory file at `~/.claude/projects/-Users-srivardhanreddygutta-automation/memory/project_opendental.md` documenting:
- This is an OpenDental desktop automation project
- Uses core/ package for automation logic
- PyWinAuto for Windows window management, PyAutoGUI for keystrokes
- CSV batch import with batch_log.csv resume
- Config in config.json with timing settings

- [ ] **Step 2: Final git status check**

Run: `git status` to verify clean working tree with all files committed.
