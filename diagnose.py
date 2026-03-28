"""Diagnostic: dump all controls pywinauto can see in OpenDental.
Run this WHILE the Select Patient panel is open."""

import sys
if sys.platform != "win32":
    print("Run this on Windows with OpenDental open!")
    sys.exit(1)

from pywinauto import Application, Desktop

print("=" * 60)
print("STEP 1: Connecting to OpenDental...")
print("=" * 60)

app = Application(backend="uia").connect(
    title_re=".*Open Dental.*|.*Demo Database.*", timeout=5
)
win = app.top_window()
print(f"Window title: {win.window_text()}")
print(f"Window rect: {win.rectangle()}")

print("\n" + "=" * 60)
print("STEP 2: ALL controls in main window (first 50)")
print("=" * 60)

count = 0
for desc in win.descendants():
    try:
        text = desc.window_text()
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        if text or auto_id:
            print(f"  [{ctrl}] text='{text}' auto_id='{auto_id}'")
            count += 1
            if count >= 50:
                print("  ... (truncated at 50)")
                break
    except Exception:
        continue

print(f"\nTotal with text/auto_id: {count}")

print("\n" + "=" * 60)
print("STEP 3: Looking for 'Add Pt' or 'Select Patient' anywhere")
print("=" * 60)

for desc in win.descendants():
    try:
        text = desc.window_text()
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        name = desc.element_info.name
        if text and ("Add" in text or "Select" in text or "Patient" in text):
            print(f"  MATCH: [{ctrl}] text='{text}' name='{name}' auto_id='{auto_id}'")
    except Exception:
        continue

print("\n" + "=" * 60)
print("STEP 4: Desktop windows search")
print("=" * 60)

desktop = Desktop(backend="uia")
for dwin in desktop.windows():
    try:
        t = dwin.window_text()
        if t and ("Open Dental" in t or "Demo" in t or "Select" in t or "Patient" in t):
            print(f"  Window: '{t}'")
    except Exception:
        continue

print("\n" + "=" * 60)
print("STEP 5: Trying win32 backend")
print("=" * 60)

try:
    app32 = Application(backend="win32").connect(
        title_re=".*Open Dental.*|.*Demo Database.*", timeout=5
    )
    win32 = app32.top_window()
    print(f"win32 window: {win32.window_text()}")
    count32 = 0
    for desc in win32.descendants():
        try:
            text = desc.window_text()
            ctrl_class = desc.friendly_class_name()
            if text and ("Add" in text or "Select" in text or "Patient" in text):
                print(f"  MATCH: [{ctrl_class}] text='{text}'")
                count32 += 1
        except Exception:
            continue
    if count32 == 0:
        print("  No matches found with win32 backend")
except Exception as e:
    print(f"  win32 backend error: {e}")

print("\nDone! Copy this output and paste it back.")
