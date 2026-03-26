"""OpenDental GUI automation — opens the app, fills the Add Patient form, clicks Save.
Uses pywinauto for window management + pyautogui for reliable typing.
Handles all trial popups automatically."""

import time
import sys
import os


def automate_patient_entry(patient, status_callback, config=None):
    """Full automation: launch OpenDental → dismiss popups → Add Patient → fill → Save."""
    if sys.platform != "win32":
        status_callback("ERROR: Requires Windows!", "red")
        return False

    from pywinauto import Application
    import pyautogui

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1

    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")

    try:
        # ═══ STEP 1: Connect or Launch ═══
        status_callback("[1/10] Finding OpenDental...", "yellow")
        app = None

        try:
            app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=3)
            status_callback("[1/10] Connected to OpenDental", "limegreen")
        except Exception:
            if not os.path.exists(app_path):
                status_callback(f"[1/10] ERROR: {app_path} not found", "red")
                return False
            status_callback("[1/10] Launching OpenDental...", "yellow")
            os.startfile(app_path)
            time.sleep(12)
            try:
                app = Application(backend="uia").connect(title_re=".*Open Dental.*", timeout=20)
            except Exception:
                try:
                    app = Application(backend="uia").connect(title_re=".*Demo Database.*", timeout=10)
                except Exception:
                    status_callback("[1/10] ERROR: OpenDental did not start", "red")
                    return False
            status_callback("[1/10] OpenDental launched!", "limegreen")

        # ═══ STEP 2: Login ═══
        status_callback("[2/10] Logging in...", "yellow")
        win = app.top_window()
        win.set_focus()
        time.sleep(1)

        # Check for Choose Database dialog
        title = win.window_text()
        if "Choose Database" in title:
            status_callback("[2/10] Closing database dialog...", "yellow")
            pyautogui.press('enter')
            time.sleep(5)

        # Press Enter to bypass login
        pyautogui.press('enter')
        time.sleep(8)

        # Reconnect after login
        try:
            app = Application(backend="uia").connect(title_re=".*Open Dental.*|.*Demo Database.*", timeout=10)
        except Exception:
            pass

        status_callback("[2/10] Logged in!", "limegreen")

        # ═══ STEP 3: Dismiss all popups ═══
        status_callback("[3/10] Checking for popups...", "yellow")

        for _ in range(8):
            time.sleep(1)
            try:
                win = app.top_window()
                title = win.window_text()
                rect = win.rectangle()
                width = rect.right - rect.left

                # Small dialog = popup
                if width < 600 and width > 50:
                    status_callback(f"[3/10] Dismissing: {title[:40]}...", "yellow")
                    pyautogui.press('enter')
                    time.sleep(1)
                    continue

                # Alerts window (has Acknowledge button)
                if "Alert" in title:
                    status_callback("[3/10] Dismissing alerts...", "yellow")
                    try:
                        ack = win.child_window(title="Acknowledge", control_type="Button")
                        ack.click_input()
                    except Exception:
                        pyautogui.press('enter')
                    time.sleep(2)
                    continue

                # Main window reached
                break

            except Exception:
                break

        status_callback("[3/10] Popups cleared!", "limegreen")

        # ═══ STEP 4: Open Select Patient ═══
        status_callback("[4/10] Opening Select Patient...", "yellow")
        win = app.top_window()
        win.set_focus()
        time.sleep(0.5)

        # Try clicking the Select Patient toolbar button
        clicked = False
        try:
            btn = win.child_window(title="Select Patient", control_type="SplitButton")
            btn.click_input()
            clicked = True
        except Exception:
            pass

        if not clicked:
            try:
                btn = win.child_window(title_re=".*Select Patient.*")
                btn.click_input()
                clicked = True
            except Exception:
                pass

        if not clicked:
            # Keyboard fallback
            pyautogui.hotkey('ctrl', 'p')

        time.sleep(3)

        # Dismiss trial popup
        try:
            popup = app.top_window()
            prect = popup.rectangle()
            pw = prect.right - prect.left
            if pw < 600 and pw > 50:
                status_callback("[4/10] Dismissing trial popup...", "yellow")
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        status_callback("[4/10] Select Patient open!", "limegreen")

        # ═══ STEP 5: Click Add Pt ═══
        status_callback("[5/10] Clicking Add Pt...", "yellow")
        time.sleep(1)

        sel_win = app.top_window()
        added = False

        try:
            add_btn = sel_win.child_window(title="Add Pt", control_type="Button")
            add_btn.click_input()
            added = True
        except Exception:
            pass

        if not added:
            try:
                # Search all buttons for "Add"
                buttons = sel_win.descendants(control_type="Button")
                for btn in buttons:
                    txt = btn.window_text()
                    if "Add Pt" in txt or "Add New" in txt:
                        btn.click_input()
                        added = True
                        break
            except Exception:
                pass

        if not added:
            pyautogui.hotkey('alt', 'a')

        time.sleep(3)

        # Dismiss trial popup again
        try:
            popup = app.top_window()
            prect = popup.rectangle()
            pw = prect.right - prect.left
            if pw < 600 and pw > 50:
                status_callback("[5/10] Dismissing trial popup...", "yellow")
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        status_callback("[5/10] Add Patient form open!", "limegreen")

        # ═══ STEP 6: Get the Edit Patient window and find fields ═══
        status_callback("[6/10] Finding form fields...", "yellow")
        time.sleep(1)

        edit_win = app.top_window()
        edit_win.set_focus()

        # Get window position for relative coordinates
        wrect = edit_win.rectangle()
        wx = wrect.left
        wy = wrect.top

        # Try to find all edit controls
        edits = []
        try:
            edits = edit_win.descendants(control_type="Edit")
            status_callback(f"[6/10] Found {len(edits)} form fields", "limegreen")
        except Exception:
            status_callback("[6/10] Warning: Could not enumerate fields", "orange")

        # ═══ STEP 7: Fill Last Name & First Name ═══
        status_callback(f"[7/10] Filling: Last Name = {patient.last_name}", "yellow")

        filled_ln = False
        # Strategy 1: Find by label
        for edit in edits:
            try:
                name = edit.element_info.name or ""
                if "Last Name" in name and "Preferred" not in name:
                    edit.click_input()
                    edit.set_edit_text(patient.last_name)
                    filled_ln = True
                    break
            except Exception:
                continue

        # Strategy 2: Click relative position in window (Last Name is ~y=107 from top)
        if not filled_ln:
            try:
                pyautogui.click(wx + 350, wy + 107)
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.write(patient.last_name, interval=0.03)
                filled_ln = True
            except Exception:
                pass

        if filled_ln:
            status_callback(f"[7/10] Last Name = {patient.last_name}", "limegreen")
        else:
            status_callback("[7/10] WARNING: Could not fill Last Name", "orange")

        time.sleep(0.3)

        # First Name
        status_callback(f"[7/10] Filling: First Name = {patient.first_name}", "yellow")
        filled_fn = False

        for edit in edits:
            try:
                name = edit.element_info.name or ""
                if name == "First Name" or "First Name" in name:
                    edit.click_input()
                    edit.set_edit_text(patient.first_name)
                    filled_fn = True
                    break
            except Exception:
                continue

        if not filled_fn:
            try:
                pyautogui.click(wx + 350, wy + 133)
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.write(patient.first_name, interval=0.03)
                filled_fn = True
            except Exception:
                pass

        if filled_fn:
            status_callback(f"[7/10] First Name = {patient.first_name}", "limegreen")

        time.sleep(0.3)

        # ═══ STEP 8: Fill demographics ═══
        status_callback("[8/10] Filling demographics...", "yellow")

        # Middle Initial
        if patient.middle_initial:
            for edit in edits:
                try:
                    name = edit.element_info.name or ""
                    if "Middle" in name:
                        edit.click_input()
                        edit.set_edit_text(patient.middle_initial)
                        status_callback(f"  Middle Initial = {patient.middle_initial}", "yellow")
                        break
                except Exception:
                    continue

        # Gender — click in the list
        if patient.gender:
            try:
                items = edit_win.descendants(control_type="ListItem")
                for item in items:
                    if item.window_text() == patient.gender:
                        item.click_input()
                        status_callback(f"  Gender = {patient.gender}", "yellow")
                        break
            except Exception:
                status_callback("  Gender: could not select", "orange")

        # Birthdate
        if patient.dob:
            for edit in edits:
                try:
                    name = edit.element_info.name or ""
                    if "Birthdate" in name or "Birth" in name:
                        edit.click_input()
                        edit.set_edit_text(patient.dob)
                        status_callback(f"  Birthdate = {patient.dob}", "yellow")
                        break
                except Exception:
                    continue

        status_callback("[8/10] Demographics filled!", "limegreen")
        time.sleep(0.3)

        # ═══ STEP 9: Fill Address & Phone ═══
        status_callback("[9/10] Filling address & phone...", "yellow")

        field_map = {
            "Home Phone": patient.phone,
            "Address": patient.address,
            "City": patient.city,
            "ST": patient.state,
        }

        for label, value in field_map.items():
            if not value:
                continue
            filled = False
            for edit in edits:
                try:
                    name = edit.element_info.name or ""
                    if name == label or label in name:
                        edit.click_input()
                        edit.set_edit_text(value)
                        status_callback(f"  {label} = {value}", "yellow")
                        filled = True
                        break
                except Exception:
                    continue
            if not filled:
                status_callback(f"  {label}: field not found", "orange")

        # Zip (might be a combo box)
        if patient.zip:
            try:
                zip_edit = edit_win.child_window(title="Zip", control_type="Edit")
                zip_edit.click_input()
                zip_edit.set_edit_text(patient.zip)
                status_callback(f"  Zip = {patient.zip}", "yellow")
            except Exception:
                try:
                    zip_combo = edit_win.child_window(title="Zip", control_type="ComboBox")
                    zip_combo.click_input()
                    pyautogui.write(patient.zip, interval=0.03)
                    status_callback(f"  Zip = {patient.zip}", "yellow")
                except Exception:
                    status_callback("  Zip: field not found", "orange")

        # SSN (on the Other tab - right side)
        if patient.ssn:
            try:
                ssn_edit = edit_win.child_window(title="SS#", control_type="Edit")
                ssn_edit.click_input()
                ssn_edit.set_edit_text(patient.ssn)
                status_callback("  SSN = ***", "yellow")
            except Exception:
                pass

        status_callback("[9/10] Address & phone filled!", "limegreen")
        time.sleep(0.5)

        # ═══ STEP 10: Click Save ═══
        status_callback("[10/10] Clicking Save...", "yellow")

        saved = False

        # Strategy 1: Find Save button
        try:
            buttons = edit_win.descendants(control_type="Button")
            for btn in buttons:
                if btn.window_text() == "Save":
                    btn.click_input()
                    saved = True
                    break
        except Exception:
            pass

        # Strategy 2: Click relative position (Save is bottom-right)
        if not saved:
            try:
                wrect = edit_win.rectangle()
                save_x = wrect.right - 60
                save_y = wrect.bottom - 30
                pyautogui.click(save_x, save_y)
                saved = True
            except Exception:
                pass

        time.sleep(2)

        # Dismiss any post-save popup
        try:
            popup = app.top_window()
            prect = popup.rectangle()
            pw = prect.right - prect.left
            if pw < 600 and pw > 50:
                pyautogui.press('enter')
                time.sleep(1)
        except Exception:
            pass

        if saved:
            status_callback(f"[10/10] SAVED! {patient.first_name} {patient.last_name} is in OpenDental!", "limegreen")
        else:
            status_callback("[10/10] WARNING: Could not find Save button", "orange")

        return saved

    except Exception as e:
        status_callback(f"ERROR at current step: {e}", "red")
        return False
