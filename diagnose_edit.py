"""Diagnostic: dump Edit Patient form controls.
Run this WHILE the Edit Patient Information window is open."""

import sys
if sys.platform != "win32":
    print("Run this on Windows with Edit Patient form open!")
    sys.exit(1)

from pywinauto import Application

print("Connecting to OpenDental...")
app = Application(backend="uia").connect(
    title_re=".*Open Dental.*|.*Demo Database.*|.*Edit Patient.*", timeout=5
)

# Try to find Edit Patient window
win = app.top_window()
title = win.window_text()
print(f"Top window: '{title}'")

# Look for Edit Patient child window
edit_win = None
try:
    edit_win = win.child_window(title_re=".*Edit Patient.*")
    if edit_win.exists(timeout=2):
        print(f"Found Edit Patient child window!")
    else:
        edit_win = None
except Exception:
    pass

if not edit_win:
    # Maybe Edit Patient IS the top window
    try:
        app2 = Application(backend="uia").connect(title_re=".*Edit Patient.*", timeout=3)
        edit_win = app2.top_window()
        print(f"Connected to Edit Patient directly: '{edit_win.window_text()}'")
    except Exception:
        print("Could not find Edit Patient window!")
        edit_win = win  # fall back to main

target = edit_win
print(f"\nScanning: '{target.window_text()}'")
print(f"Rect: {target.rectangle()}")

print("\n--- EDIT FIELDS (auto_id containing 'text' or type=Edit) ---\n")
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        name = desc.element_info.name
        if ctrl == "Edit" or (auto_id and "text" in auto_id.lower()):
            r = ""
            try:
                r = desc.rectangle()
            except Exception:
                pass
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' name='{name}' rect={r}")
    except Exception:
        continue

print("\n--- BUTTONS (auto_id containing 'but' or type has 'Button'/'Custom') ---\n")
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        if auto_id and ("but" in auto_id.lower() or "save" in auto_id.lower()):
            r = ""
            try:
                r = desc.rectangle()
            except Exception:
                pass
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' rect={r}")
        elif "Save" in text or "OK" in text:
            r = ""
            try:
                r = desc.rectangle()
            except Exception:
                pass
            print(f"[{ctrl}] auto_id='{auto_id}' text='{text}' rect={r}")
    except Exception:
        continue

print("\n--- ALL CONTROLS WITH LABELS (name has content) ---\n")
count = 0
for desc in target.descendants():
    try:
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        text = desc.window_text()
        name = desc.element_info.name
        if name and name.strip():
            print(f"[{ctrl}] auto_id='{auto_id}' name='{name}' text='{text}'")
            count += 1
            if count >= 80:
                print("... (truncated)")
                break
    except Exception:
        continue

print(f"\nTotal labeled: {count}")
print("\nDone! Paste this output back.")
