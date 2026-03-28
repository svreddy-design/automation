"""Diagnostic: dump Edit Patient form controls.
The form is embedded inside the main window as a tab/panel."""

import sys
if sys.platform != "win32":
    print("Run this on Windows!")
    sys.exit(1)

from pywinauto import Application

print("Connecting to OpenDental main window...")
app = Application(backend="uia").connect(
    title_re=".*Open Dental.*|.*Demo Database.*", timeout=5
)
win = app.top_window()
print(f"Window: '{win.window_text()}'")

# Look for Edit Patient child form
print("\nSearching for Edit Patient form inside main window...")
edit_form = None
for desc in win.descendants():
    try:
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        if "Edit Patient" in text or "FormPatientEdit" in (auto_id or ""):
            print(f"  FOUND: [{desc.element_info.control_type}] auto_id='{auto_id}' text='{text}'")
            if not edit_form:
                edit_form = desc
    except Exception:
        continue

# Scan for ALL Edit fields, using main window
target = win
print(f"\nScanning ALL Edit fields in main window...\n")

print("--- EDIT FIELDS ---\n")
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id or ""
        text = desc.window_text()
        name = desc.element_info.name or ""
        if ctrl == "Edit":
            r = desc.rectangle()
            print(f"[Edit] auto_id='{auto_id}' text='{text}' name='{name}' rect={r}")
    except Exception:
        continue

print("\n--- BUTTONS WITH 'Save' or 'OK' or 'but' ---\n")
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id or ""
        text = desc.window_text()
        if "Save" in text or "OK" in text or "save" in auto_id.lower() or "butOK" in auto_id:
            r = desc.rectangle()
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' rect={r}")
    except Exception:
        continue

print("\n--- LABELS (Text controls near edit fields) ---\n")
count = 0
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id or ""
        text = desc.window_text()
        name = desc.element_info.name or ""
        # Only show labels relevant to patient form
        if ctrl == "Text" and name and any(kw in name for kw in [
            "Name", "Phone", "Address", "City", "State", "Zip", "Birth",
            "SSN", "Gender", "Save", "Employer", "Email", "Medicaid",
            "Patient Number", "Salutation", "Title", "Preferred", "ST"
        ]):
            print(f"[Text] auto_id='{auto_id}' name='{name}'")
            count += 1
    except Exception:
        continue

print(f"\nTotal labels: {count}")
print("\nDone!")
