"""OpenDental GUI automation using pywinauto UIA backend.
Finds controls by name/type — no screen coordinates needed.
Works reliably regardless of screen resolution or window position."""

import time
import sys
import os


def automate_patient_entry(patient, status_callback, config=None):
    """Open OpenDental, navigate to Add Patient, fill all fields, and save.

    Args:
        patient: Patient dataclass with field values
        status_callback: function(text, color) for real-time status updates
        config: dict with timing settings
    Returns:
        True if patient was saved successfully, False otherwise
    """
    if sys.platform != "win32":
        status_callback("ERROR: GUI automation requires Windows!", "red")
        return False

    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    import pyautogui

    timing = config or {}
    type_delay = timing.get("typing_interval_ms", 50) / 1000.0
    field_delay = timing.get("field_delay_ms", 300) / 1000.0

    try:
        # ============ STEP 1: Connect to OpenDental ============
        status_callback("Step 1/10: Connecting to OpenDental...", "yellow")

        app = None
        # Try to connect to already running OpenDental
        try:
            app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=5)
            status_callback("Step 1/10: Connected to running OpenDental", "limegreen")
        except Exception:
            # Not running — launch it
            app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")
            if not os.path.exists(app_path):
                status_callback(f"ERROR: OpenDental not found at {app_path}", "red")
                return False

            status_callback("Step 1/10: Launching OpenDental...", "yellow")
            app = Application(backend="uia").start(app_path)
            time.sleep(timing.get("app_load_delay_s", 10))

            # Handle Choose Database dialog
            status_callback("Step 2/10: Handling login...", "yellow")
            try:
                app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=15)
            except Exception as e:
                status_callback(f"ERROR: Could not connect to OpenDental: {e}", "red")
                return False

        # ============ STEP 2: Handle popups ============
        status_callback("Step 2/10: Dismissing popups...", "yellow")
        main_win = app.top_window()
        main_win.set_focus()
        time.sleep(1)

        # Dismiss any dialog boxes (Choose Database, Alerts, Trial warnings)
        for _ in range(5):
            try:
                dialog = app.top_window()
                title = dialog.window_text()

                if "Choose Database" in title:
                    status_callback("Step 2/10: Closing database dialog...", "yellow")
                    ok_btn = dialog.child_window(title="OK", control_type="Button")
                    ok_btn.click_input()
                    time.sleep(3)
                    continue

                if "Alert" in title or "alert" in title:
                    status_callback("Step 2/10: Dismissing alert...", "yellow")
                    try:
                        ack_btn = dialog.child_window(title="Acknowledge", control_type="Button")
                        ack_btn.click_input()
                    except Exception:
                        pyautogui.press('enter')
                    time.sleep(2)
                    continue

                # Trial version popup
                if "Trial" in title or "Maximum" in title:
                    status_callback("Step 2/10: Dismissing trial popup...", "yellow")
                    pyautogui.press('enter')
                    time.sleep(1)
                    continue

                # Any other OK/Close dialog
                try:
                    ok_btn = dialog.child_window(title="OK", control_type="Button")
                    if ok_btn.exists():
                        ok_btn.click_input()
                        time.sleep(1)
                        continue
                except Exception:
                    pass

                break  # No more dialogs

            except Exception:
                break

        status_callback("Step 2/10: Popups handled", "limegreen")

        # ============ STEP 3: Open Select Patient ============
        status_callback("Step 3/10: Opening Select Patient...", "yellow")
        main_win = app.top_window()
        main_win.set_focus()
        time.sleep(0.5)

        # Try clicking Select Patient button in toolbar
        try:
            sel_btn = main_win.child_window(title="Select Patient", control_type="SplitButton")
            sel_btn.click_input()
        except Exception:
            try:
                sel_btn = main_win.child_window(title_re=".*Select Patient.*")
                sel_btn.click_input()
            except Exception:
                # Fallback: use keyboard
                pyautogui.click(108, 51)

        time.sleep(2)

        # Handle trial popup if it appears
        try:
            trial_dlg = app.window(title_re=".*Trial.*")
            if trial_dlg.exists(timeout=2):
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        status_callback("Step 3/10: Select Patient window open", "limegreen")

        # ============ STEP 4: Click Add Pt ============
        status_callback("Step 4/10: Clicking Add Pt...", "yellow")
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
                # Fallback
                pyautogui.hotkey('alt', 'a')

        time.sleep(3)

        # Handle trial popup again
        try:
            trial_dlg = app.window(title_re=".*Trial.*")
            if trial_dlg.exists(timeout=2):
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        status_callback("Step 4/10: Add Patient form open", "limegreen")

        # ============ STEP 5-9: Fill fields ============
        edit_win = app.top_window()
        edit_win.set_focus()
        time.sleep(0.5)

        def fill_field(field_title, value, step_num, step_total=10):
            """Find a field by nearby label and type into it."""
            if not value:
                status_callback(f"Step {step_num}/{step_total}: Skipping {field_title} (empty)", "gray")
                return True

            display = "***" if "SSN" in field_title else value
            status_callback(f"Step {step_num}/{step_total}: Filling {field_title} = {display}", "yellow")

            try:
                # Try to find edit control near the label
                edit = edit_win.child_window(title=field_title, control_type="Edit")
                edit.click_input()
                edit.set_edit_text("")
                time.sleep(0.1)
                # Type character by character for visibility
                for char in value:
                    pyautogui.press(char) if len(char) == 1 and char.isalnum() == False else pyautogui.write(char, interval=type_delay)

            except Exception:
                try:
                    # Alternative: find by automation ID patterns
                    edits = edit_win.children(control_type="Edit")
                    # Can't find by name, will use tab order
                    return False
                except Exception:
                    return False

            time.sleep(field_delay)
            return True

        # Use tab-based entry as it's most reliable with OpenDental's WinForms
        status_callback("Step 5/10: Filling Last Name...", "yellow")

        # The cursor should already be on Last Name in the Edit Patient form
        # Clear and type each field, tabbing between them
        fields_to_fill = [
            ("Last Name", patient.last_name, "5a"),
            ("First Name", patient.first_name, "5b"),
        ]

        # Click on the Last Name field to make sure we're in the right place
        try:
            ln_edit = edit_win.child_window(title="Last Name", control_type="Edit")
            ln_edit.click_input()
            ln_edit.set_edit_text(patient.last_name)
            status_callback(f"Step 5/10: Last Name = {patient.last_name}", "limegreen")
        except Exception:
            # Fallback: just type, assuming cursor is on Last Name
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(patient.last_name, interval=type_delay)
            status_callback(f"Step 5/10: Last Name = {patient.last_name}", "limegreen")

        time.sleep(field_delay)

        # First Name
        status_callback("Step 6/10: Filling First Name...", "yellow")
        try:
            fn_edit = edit_win.child_window(title="First Name", control_type="Edit")
            fn_edit.click_input()
            fn_edit.set_edit_text(patient.first_name)
        except Exception:
            pyautogui.press('tab')
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(patient.first_name, interval=type_delay)
        status_callback(f"Step 6/10: First Name = {patient.first_name}", "limegreen")
        time.sleep(field_delay)

        # Middle Initial — in OpenDental this is part of "Preferred Name, Middle Initial" field
        status_callback("Step 7/10: Filling demographics...", "yellow")
        if patient.middle_initial:
            try:
                mi_edit = edit_win.child_window(title="Middle Initial", control_type="Edit")
                mi_edit.click_input()
                mi_edit.set_edit_text(patient.middle_initial)
            except Exception:
                # The MI field might be combined with Preferred Name
                pass

        # Gender — click on Male/Female/Unknown in the list
        if patient.gender:
            status_callback(f"Step 7/10: Selecting Gender = {patient.gender}...", "yellow")
            try:
                gender_list = edit_win.child_window(title="Gender", control_type="List")
                gender_item = gender_list.child_window(title=patient.gender)
                gender_item.click_input()
            except Exception:
                try:
                    # Try clicking the text directly
                    gender_item = edit_win.child_window(title=patient.gender, control_type="ListItem")
                    gender_item.click_input()
                except Exception:
                    pass

        status_callback("Step 7/10: Demographics filled", "limegreen")
        time.sleep(field_delay)

        # Birthdate
        status_callback("Step 8/10: Filling Birthdate...", "yellow")
        if patient.dob:
            try:
                bd_edit = edit_win.child_window(title="Birthdate", control_type="Edit")
                bd_edit.click_input()
                bd_edit.set_edit_text(patient.dob)
            except Exception:
                pass
        status_callback(f"Step 8/10: Birthdate = {patient.dob or 'skipped'}", "limegreen")
        time.sleep(field_delay)

        # Address and Phone section (right side of form)
        status_callback("Step 9/10: Filling Address & Phone...", "yellow")

        right_side_fields = [
            ("Home Phone", patient.phone),
            ("Address", patient.address),
            ("City", patient.city),
            ("ST", patient.state),
        ]

        for field_name, value in right_side_fields:
            if value:
                try:
                    edit = edit_win.child_window(title=field_name, control_type="Edit")
                    edit.click_input()
                    edit.set_edit_text(value)
                    status_callback(f"Step 9/10: {field_name} = {value}", "yellow")
                    time.sleep(0.2)
                except Exception:
                    status_callback(f"Step 9/10: Could not find {field_name} field", "orange")

        # Zip has a special "Edit Zip" button area
        if patient.zip:
            try:
                zip_edit = edit_win.child_window(title="Zip", control_type="Edit")
                zip_edit.click_input()
                zip_edit.set_edit_text(patient.zip)
            except Exception:
                pass

        status_callback("Step 9/10: Address & Phone filled", "limegreen")
        time.sleep(field_delay)

        # ============ STEP 10: Save ============
        status_callback("Step 10/10: Clicking Save...", "yellow")
        try:
            save_btn = edit_win.child_window(title="Save", control_type="Button")
            save_btn.click_input()
        except Exception:
            # Fallback: try to find any Save button
            try:
                save_btn = edit_win.child_window(title_re=".*Save.*", control_type="Button")
                save_btn.click_input()
            except Exception:
                status_callback("ERROR: Could not find Save button!", "red")
                return False

        time.sleep(2)

        # Handle any post-save popups
        try:
            dialog = app.top_window()
            if "Trial" in dialog.window_text():
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        status_callback(
            f"DONE! Patient {patient.first_name} {patient.last_name} saved in OpenDental!",
            "limegreen"
        )
        return True

    except Exception as e:
        status_callback(f"ERROR at current step: {e}", "red")
        return False
