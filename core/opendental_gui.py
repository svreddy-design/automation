"""OpenDental GUI automation using pywinauto.
Handles all popups, finds controls by name, fills every field, clicks Save."""

import time
import sys
import os


def dismiss_all_popups(app, status_callback):
    """Keep dismissing popups until we reach the main OpenDental window."""
    for attempt in range(10):
        try:
            time.sleep(1)
            win = app.top_window()
            title = win.window_text()
            status_callback(f"  Popup: '{title}'", "yellow")

            # Choose Database dialog
            if "Choose Database" in title:
                status_callback("  Dismissing: Choose Database...", "yellow")
                try:
                    win.child_window(title="OK", control_type="Button").click_input()
                except Exception:
                    win.type_keys("{ENTER}")
                time.sleep(3)
                continue

            # Alerts dialog
            if "Alerts" in title or "Alert" in title:
                status_callback("  Dismissing: Alerts...", "yellow")
                try:
                    win.child_window(title="Acknowledge", control_type="Button").click_input()
                except Exception:
                    try:
                        win.child_window(title_re=".*Acknowledge.*").click_input()
                    except Exception:
                        win.type_keys("{ENTER}")
                time.sleep(2)
                continue

            # Trial version popup
            if "Trial" in title:
                status_callback("  Dismissing: Trial warning...", "yellow")
                try:
                    win.child_window(title="OK", control_type="Button").click_input()
                except Exception:
                    win.type_keys("{ENTER}")
                time.sleep(1)
                continue

            # Generic OK dialog
            try:
                ok_btn = win.child_window(title="OK", control_type="Button")
                if ok_btn.exists(timeout=1):
                    # Check if this is a small dialog (not the main window)
                    if win.rectangle().width() < 600:
                        status_callback(f"  Dismissing: {title}...", "yellow")
                        ok_btn.click_input()
                        time.sleep(1)
                        continue
            except Exception:
                pass

            # If we get here, it's probably the main window
            if "Open Dental" in title or "Demo Database" in title:
                status_callback("  Main window reached!", "limegreen")
                return True

        except Exception:
            continue

    return True


def automate_patient_entry(patient, status_callback, config=None):
    """Open OpenDental, navigate to Add Patient, fill all fields, and save."""
    if sys.platform != "win32":
        status_callback("ERROR: GUI automation requires Windows!", "red")
        return False

    from pywinauto import Application
    import pyautogui

    timing = config or {}
    type_delay = timing.get("typing_interval_ms", 50) / 1000.0

    try:
        # ========== STEP 1: Connect to OpenDental ==========
        status_callback("Step 1/10: Finding OpenDental...", "yellow")
        app = None

        # Try connecting to already-running OpenDental
        try:
            app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=3)
            status_callback("Step 1/10: Connected to running OpenDental!", "limegreen")
        except Exception:
            # Launch it
            app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")
            if not os.path.exists(app_path):
                status_callback(f"ERROR: Not found: {app_path}", "red")
                return False

            status_callback("Step 1/10: Launching OpenDental...", "yellow")
            Application(backend="uia").start(app_path)
            time.sleep(timing.get("app_load_delay_s", 12))

            try:
                app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=15)
            except Exception:
                try:
                    app = Application(backend="uia").connect(title_re=".*Demo Database.*", timeout=10)
                except Exception as e:
                    status_callback(f"ERROR: Cannot connect: {e}", "red")
                    return False

            status_callback("Step 1/10: OpenDental launched!", "limegreen")

        # ========== STEP 2: Handle login ==========
        status_callback("Step 2/10: Handling login...", "yellow")
        time.sleep(2)

        # Check if we're at the login screen (just press Enter)
        try:
            win = app.top_window()
            title = win.window_text()
            if "Open Dental" in title and "Admin" not in title:
                # Might be at login - press Enter
                win.type_keys("{ENTER}")
                time.sleep(timing.get("login_delay_s", 8))
        except Exception:
            pass

        status_callback("Step 2/10: Login done!", "limegreen")

        # ========== STEP 3: Dismiss all popups ==========
        status_callback("Step 3/10: Dismissing popups...", "yellow")

        # Re-connect after login (window title changes)
        try:
            app = Application(backend="uia").connect(title_re=".*Demo Database.*", timeout=10)
        except Exception:
            try:
                app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=10)
            except Exception:
                pass

        dismiss_all_popups(app, status_callback)
        status_callback("Step 3/10: All popups dismissed!", "limegreen")

        # ========== STEP 4: Open Select Patient ==========
        status_callback("Step 4/10: Opening Select Patient...", "yellow")
        main_win = app.top_window()
        main_win.set_focus()
        time.sleep(0.5)

        # Click Select Patient in toolbar
        try:
            sel_btn = main_win.child_window(title="Select Patient", control_type="SplitButton")
            sel_btn.click_input()
        except Exception:
            try:
                sel_btn = main_win.child_window(title_re=".*Select Patient.*")
                sel_btn.click_input()
            except Exception:
                # Use keyboard shortcut
                main_win.type_keys("^p")  # Ctrl+P

        time.sleep(3)

        # Dismiss trial popup if it appeared
        try:
            trial_win = app.top_window()
            if trial_win.rectangle().width() < 500:
                trial_win.type_keys("{ENTER}")
                time.sleep(1)
        except Exception:
            pass

        status_callback("Step 4/10: Select Patient open!", "limegreen")

        # ========== STEP 5: Click Add Pt ==========
        status_callback("Step 5/10: Clicking Add Pt...", "yellow")
        time.sleep(1)

        sel_win = app.top_window()
        try:
            add_btn = sel_win.child_window(title="Add Pt", control_type="Button")
            add_btn.click_input()
        except Exception:
            try:
                add_btn = sel_win.child_window(title_re=".*Add Pt.*")
                add_btn.click_input()
            except Exception:
                try:
                    add_btn = sel_win.child_window(title="Add New Family:", control_type="Button")
                    add_btn.click_input()
                except Exception:
                    sel_win.type_keys("%a")  # Alt+A

        time.sleep(3)

        # Dismiss trial popup again if needed
        try:
            trial_win = app.top_window()
            title = trial_win.window_text()
            if trial_win.rectangle().width() < 500 or "Trial" in title:
                trial_win.type_keys("{ENTER}")
                time.sleep(1)
        except Exception:
            pass

        status_callback("Step 5/10: Add Patient form open!", "limegreen")

        # ========== STEP 6: Fill Last Name ==========
        status_callback("Step 6/10: Filling Last Name...", "yellow")
        time.sleep(1)
        edit_win = app.top_window()
        edit_win.set_focus()

        # Try multiple ways to find and fill Last Name
        filled_ln = False
        try:
            ln_edit = edit_win.child_window(title="Last Name", control_type="Edit")
            ln_edit.click_input()
            ln_edit.set_edit_text(patient.last_name)
            filled_ln = True
        except Exception:
            pass

        if not filled_ln:
            try:
                # Get all edit controls and fill the first one (Last Name is first)
                edits = edit_win.children(control_type="Edit")
                if len(edits) > 0:
                    edits[0].click_input()
                    edits[0].set_edit_text(patient.last_name)
                    filled_ln = True
            except Exception:
                pass

        if not filled_ln:
            # Last resort: just type
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(patient.last_name, interval=type_delay)
            filled_ln = True

        status_callback(f"Step 6/10: Last Name = {patient.last_name}", "limegreen")
        time.sleep(0.3)

        # ========== STEP 7: Fill First Name ==========
        status_callback("Step 7/10: Filling First Name...", "yellow")
        filled_fn = False
        try:
            fn_edit = edit_win.child_window(title="First Name", control_type="Edit")
            fn_edit.click_input()
            fn_edit.set_edit_text(patient.first_name)
            filled_fn = True
        except Exception:
            pass

        if not filled_fn:
            try:
                edits = edit_win.children(control_type="Edit")
                if len(edits) > 1:
                    edits[1].click_input()
                    edits[1].set_edit_text(patient.first_name)
                    filled_fn = True
            except Exception:
                pass

        if not filled_fn:
            pyautogui.press('tab')
            pyautogui.write(patient.first_name, interval=type_delay)

        status_callback(f"Step 7/10: First Name = {patient.first_name}", "limegreen")
        time.sleep(0.3)

        # ========== STEP 8: Fill remaining demographics ==========
        status_callback("Step 8/10: Filling demographics...", "yellow")

        # Middle Initial
        if patient.middle_initial:
            try:
                mi_edit = edit_win.child_window(title="Middle Initial", control_type="Edit")
                mi_edit.click_input()
                mi_edit.set_edit_text(patient.middle_initial)
                status_callback(f"  Middle Initial = {patient.middle_initial}", "yellow")
            except Exception:
                pass

        # Gender - click in the Gender list
        if patient.gender:
            try:
                gender_item = edit_win.child_window(title=patient.gender, control_type="ListItem")
                gender_item.click_input()
                status_callback(f"  Gender = {patient.gender}", "yellow")
            except Exception:
                try:
                    # Try finding "Male" or "Female" text directly
                    gender_item = edit_win.child_window(title=patient.gender)
                    gender_item.click_input()
                except Exception:
                    pass

        # Birthdate
        if patient.dob:
            try:
                bd_edit = edit_win.child_window(title="Birthdate", control_type="Edit")
                bd_edit.click_input()
                bd_edit.set_edit_text(patient.dob)
                status_callback(f"  Birthdate = {patient.dob}", "yellow")
            except Exception:
                pass

        # SSN
        if patient.ssn:
            try:
                ssn_edit = edit_win.child_window(title="SSN", control_type="Edit")
                ssn_edit.click_input()
                ssn_edit.set_edit_text(patient.ssn)
                status_callback("  SSN = ***", "yellow")
            except Exception:
                pass

        status_callback("Step 8/10: Demographics filled!", "limegreen")
        time.sleep(0.3)

        # ========== STEP 9: Fill Address & Phone ==========
        status_callback("Step 9/10: Filling Address & Phone...", "yellow")

        address_fields = {
            "Home Phone": patient.phone,
            "Address": patient.address,
            "City": patient.city,
            "ST": patient.state,
            "Zip": patient.zip,
        }

        for field_name, value in address_fields.items():
            if not value:
                continue
            try:
                edit = edit_win.child_window(title=field_name, control_type="Edit")
                edit.click_input()
                edit.set_edit_text(value)
                status_callback(f"  {field_name} = {value}", "yellow")
                time.sleep(0.2)
            except Exception:
                status_callback(f"  Warning: Could not find '{field_name}' field", "orange")

        status_callback("Step 9/10: Address & Phone filled!", "limegreen")
        time.sleep(0.3)

        # ========== STEP 10: Click Save ==========
        status_callback("Step 10/10: Saving patient...", "yellow")

        saved = False
        try:
            save_btn = edit_win.child_window(title="Save", control_type="Button")
            save_btn.click_input()
            saved = True
        except Exception:
            pass

        if not saved:
            try:
                # Find button with Save in the text
                buttons = edit_win.children(control_type="Button")
                for btn in buttons:
                    if "Save" in btn.window_text():
                        btn.click_input()
                        saved = True
                        break
            except Exception:
                pass

        if not saved:
            # Last resort: Enter key
            edit_win.type_keys("{ENTER}")
            saved = True

        time.sleep(2)

        # Dismiss any post-save popups
        try:
            post_win = app.top_window()
            if post_win.rectangle().width() < 500:
                post_win.type_keys("{ENTER}")
                time.sleep(1)
        except Exception:
            pass

        status_callback(
            f"DONE! {patient.first_name} {patient.last_name} saved in OpenDental!",
            "limegreen"
        )
        return True

    except Exception as e:
        status_callback(f"ERROR: {e}", "red")
        return False
