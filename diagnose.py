"""Diagnostic: dump all controls pywinauto can see in OpenDental.

Run this at EACH step of the manual workflow:
  Step 1: Main screen (before clicking anything)
  Step 2: After clicking Select Patient
  Step 3: After clicking Add Pt (on the Edit Patient form)

Usage: python diagnose.py > step1.txt
       python diagnose.py > step2.txt
       python diagnose.py > step3.txt
"""

import sys
if sys.platform != "win32":
    print("Run this on Windows with OpenDental open!")
    sys.exit(1)

from pywinauto import Application

print("Connecting to OpenDental...")
app = Application(backend="uia").connect(
    title_re=".*Open Dental.*|.*Demo Database.*", timeout=5
)
win = app.top_window()
title = win.window_text()
rect = win.rectangle()
print(f"Window: '{title}'")
print(f"Rect: {rect}")

print("\n--- ALL CONTROLS (with text or auto_id) ---\n")

count = 0
for desc in win.descendants():
    try:
        text = desc.window_text()
        ctrl = desc.element_info.control_type
        auto_id = desc.element_info.automation_id
        name = desc.element_info.name
        if text or auto_id:
            r = ""
            try:
                r = desc.rectangle()
            except Exception:
                pass
            print(f"[{ctrl}] text='{text}' auto_id='{auto_id}' name='{name}' rect={r}")
            count += 1
    except Exception:
        continue

print(f"\nTotal: {count} controls")
print("\nDone!")
