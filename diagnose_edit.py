"""Diagnostic: dump Edit Patient Information controls.
Run this WHILE the Edit Patient Information window is open."""

import sys
if sys.platform != "win32":
    print("Run this on Windows!")
    sys.exit(1)

from pywinauto import Application

# Connect directly to "Edit Patient Information" window
print("Connecting to Edit Patient Information...")
try:
    app = Application(backend="uia").connect(title_re=".*Edit Patient.*", timeout=5)
    win = app.top_window()
except Exception:
    print("Could not find 'Edit Patient Information' window!")
    print("Make sure the Edit Patient form is open in OpenDental.")
    sys.exit(1)

title = win.window_text()
rect = win.rectangle()
print(f"Window: '{title}'")
print(f"Rect: {rect}")

print("\n--- EDIT FIELDS ---\n")
for desc in win.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        name = desc.element_info.name
        if ctrl == "Edit":
            r = desc.rectangle()
            print(f"[Edit] auto_id='{auto_id}' text='{text}' name='{name}' rect={r}")
    except Exception:
        continue

print("\n--- BUTTONS ---\n")
for desc in win.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        if "Save" in text or "OK" in text or (auto_id and "but" in auto_id.lower()):
            r = desc.rectangle()
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' rect={r}")
    except Exception:
        continue

print("\n--- DROPDOWNS / COMBOS ---\n")
for desc in win.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        name = desc.element_info.name
        if ctrl in ("ComboBox", "List") or (auto_id and "combo" in auto_id.lower()):
            r = desc.rectangle()
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' name='{name}' rect={r}")
    except Exception:
        continue

print("\nDone!")
