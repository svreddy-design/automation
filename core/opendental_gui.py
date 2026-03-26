"""Smart OpenDental GUI automation.
The bot checks what window is showing and decides what to do.
Handles every popup, dialog, and screen intelligently."""

import time
import sys
import os


def what_am_i_looking_at(app):
    """Check the current top window and return what screen we're on."""
    try:
        win = app.top_window()
        title = win.window_text()
        rect = win.rectangle()
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        if "Choose Database" in title:
            return "choose_database", win
        if "Alert" in title:
            return "alerts", win
        if "Trial" in title:
            return "trial_popup", win
        if "Select Patient" in title:
            return "select_patient", win
        if "Edit Patient" in title:
            return "edit_patient", win
        if "Open Dental" in title or "Demo Database" in title:
            # Could be main window OR a small popup on top
            if width < 500 and height < 300:
                return "small_popup", win
            return "main_window", win

        # Unknown small dialog
        if width < 600 and height < 400:
            return "unknown_popup", win

        return "unknown", win

    except Exception:
        return "error", None


def handle_screen(screen_type, win, patient, status_callback, app, edits_cache):
    """Take action based on what screen we're looking at."""
    import pyautogui

    if screen_type == "choose_database":
        status_callback("  [BRAIN] See: Choose Database → clicking OK", "cyan")
        try:
            ok_btn = win.child_window(title="OK", control_type="Button")
            ok_btn.click_input()
        except Exception:
            pyautogui.press('enter')
        time.sleep(5)
        return "continue"

    elif screen_type == "alerts":
        status_callback("  [BRAIN] See: Alerts → clicking Acknowledge", "cyan")
        try:
            ack = win.child_window(title="Acknowledge", control_type="Button")
            ack.click_input()
        except Exception:
            pyautogui.press('enter')
        time.sleep(2)
        return "continue"

    elif screen_type in ("trial_popup", "small_popup", "unknown_popup"):
        status_callback(f"  [BRAIN] See: Popup ({win.window_text()[:30]}) → dismissing", "cyan")
        pyautogui.press('enter')
        time.sleep(1)
        return "continue"

    elif screen_type == "main_window":
        return "at_main"

    elif screen_type == "select_patient":
        return "at_select_patient"

    elif screen_type == "edit_patient":
        return "at_edit_patient"

    else:
        status_callback(f"  [BRAIN] See: Unknown window ({win.window_text()[:30]})", "orange")
        pyautogui.press('enter')
        time.sleep(1)
        return "continue"


def fill_patient_form(win, patient, status_callback):
    """Fill all fields in the Edit Patient Information form."""
    import pyautogui

    # Get all edit controls in the form
    edits = []
    try:
        edits = win.descendants(control_type="Edit")
        status_callback(f"  Found {len(edits)} fields in form", "cyan")
    except Exception:
        status_callback("  WARNING: Could not find form fields", "orange")

    def fill_by_name(name, value, display_name=None):
        """Try to fill a field by its automation name."""
        if not value:
            return True
        disp = "***" if "SS" in name else value
        if display_name:
            disp_name = display_name
        else:
            disp_name = name

        # Strategy 1: Search edits by name
        for edit in edits:
            try:
                ename = edit.element_info.name or ""
                if name in ename:
                    edit.click_input()
                    time.sleep(0.1)
                    edit.set_edit_text(value)
                    status_callback(f"  {disp_name} = {disp}", "yellow")
                    return True
            except Exception:
                continue

        # Strategy 2: Try child_window
        try:
            edit = win.child_window(title=name, control_type="Edit")
            edit.click_input()
            time.sleep(0.1)
            edit.set_edit_text(value)
            status_callback(f"  {disp_name} = {disp}", "yellow")
            return True
        except Exception:
            pass

        status_callback(f"  {disp_name}: field not found (skipped)", "orange")
        return False

    # Fill each field
    fill_by_name("Last Name", patient.last_name)
    time.sleep(0.2)
    fill_by_name("First Name", patient.first_name)
    time.sleep(0.2)

    if patient.middle_initial:
        fill_by_name("Middle Initial", patient.middle_initial, "Middle Initial")
        time.sleep(0.2)

    # Gender — click in list
    if patient.gender:
        try:
            items = win.descendants(control_type="ListItem")
            for item in items:
                if item.window_text() == patient.gender:
                    item.click_input()
                    status_callback(f"  Gender = {patient.gender}", "yellow")
                    break
        except Exception:
            status_callback("  Gender: could not select", "orange")

    if patient.dob:
        fill_by_name("Birthdate", patient.dob)
        time.sleep(0.2)

    # Right side fields
    fill_by_name("Home Phone", patient.phone)
    time.sleep(0.2)
    fill_by_name("Address", patient.address)
    time.sleep(0.2)
    fill_by_name("City", patient.city)
    time.sleep(0.2)
    fill_by_name("ST", patient.state)
    time.sleep(0.2)

    # Zip might be combo
    if patient.zip:
        filled = fill_by_name("Zip", patient.zip)
        if not filled:
            try:
                combo = win.child_window(title="Zip", control_type="ComboBox")
                combo.click_input()
                pyautogui.write(patient.zip, interval=0.03)
                status_callback(f"  Zip = {patient.zip}", "yellow")
            except Exception:
                pass

    if patient.ssn:
        fill_by_name("SS#", patient.ssn, "SSN")

    return True


def click_save(win, status_callback):
    """Find and click the Save button."""
    import pyautogui

    # Strategy 1: Find by name
    try:
        buttons = win.descendants(control_type="Button")
        for btn in buttons:
            if btn.window_text() == "Save":
                btn.click_input()
                return True
    except Exception:
        pass

    # Strategy 2: Bottom-right corner of window
    try:
        rect = win.rectangle()
        pyautogui.click(rect.right - 60, rect.bottom - 30)
        return True
    except Exception:
        pass

    status_callback("  WARNING: Could not find Save button", "orange")
    return False


def automate_patient_entry(patient, status_callback, config=None):
    """Smart automation: checks what's on screen, decides what to do."""
    if sys.platform != "win32":
        status_callback("ERROR: Requires Windows!", "red")
        return False

    from pywinauto import Application
    import pyautogui

    pyautogui.FAILSAFE = True
    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")

    try:
        # ═══ PHASE 1: Get OpenDental Running ═══
        status_callback("[PHASE 1] Getting OpenDental running...", "yellow")

        app = None
        try:
            app = Application(backend="uia").connect(title_re=".*Open Dental.*|.*Demo Database.*", timeout=3)
            status_callback("[PHASE 1] OpenDental is running!", "limegreen")
        except Exception:
            if not os.path.exists(app_path):
                status_callback(f"ERROR: {app_path} not found", "red")
                return False
            status_callback("[PHASE 1] Launching OpenDental...", "yellow")
            os.startfile(app_path)
            time.sleep(12)
            try:
                app = Application(backend="uia").connect(title_re=".*Open Dental.*|.*Demo Database.*", timeout=20)
            except Exception:
                status_callback("ERROR: OpenDental failed to start", "red")
                return False
            status_callback("[PHASE 1] OpenDental launched!", "limegreen")

        # ═══ PHASE 2: Navigate to Main Window ═══
        status_callback("[PHASE 2] Navigating to main screen...", "yellow")

        # Keep handling whatever we see until we reach the main window
        max_attempts = 15
        for attempt in range(max_attempts):
            screen, win = what_am_i_looking_at(app)
            status_callback(f"  [BRAIN] Screen: {screen} (attempt {attempt+1})", "cyan")

            result = handle_screen(screen, win, patient, status_callback, app, None)

            if result == "at_main":
                status_callback("[PHASE 2] At main window!", "limegreen")
                break
            elif result == "at_select_patient":
                status_callback("[PHASE 2] Already at Select Patient!", "limegreen")
                break
            elif result == "at_edit_patient":
                status_callback("[PHASE 2] Already at Edit Patient!", "limegreen")
                break

            # After handling, reconnect in case window changed
            try:
                app = Application(backend="uia").connect(
                    title_re=".*Open Dental.*|.*Demo Database.*|.*Select Patient.*|.*Edit Patient.*",
                    timeout=5
                )
            except Exception:
                pass

        # ═══ PHASE 3: Open Select Patient ═══
        screen, win = what_am_i_looking_at(app)

        if screen == "at_edit_patient" or screen == "edit_patient":
            status_callback("[PHASE 3] Already on Edit Patient form!", "limegreen")
        elif screen == "at_select_patient" or screen == "select_patient":
            status_callback("[PHASE 3] Already on Select Patient — clicking Add Pt...", "yellow")
        else:
            status_callback("[PHASE 3] Opening Select Patient...", "yellow")
            win = app.top_window()
            win.set_focus()

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
                pyautogui.hotkey('ctrl', 'p')

            time.sleep(3)

            # Handle any popup that appeared
            screen, win = what_am_i_looking_at(app)
            while screen in ("trial_popup", "small_popup", "unknown_popup"):
                handle_screen(screen, win, patient, status_callback, app, None)
                time.sleep(1)
                try:
                    app = Application(backend="uia").connect(
                        title_re=".*Select Patient.*|.*Open Dental.*",
                        timeout=5
                    )
                except Exception:
                    pass
                screen, win = what_am_i_looking_at(app)

            status_callback("[PHASE 3] Select Patient open!", "limegreen")

        # ═══ PHASE 4: Click Add Pt ═══
        screen, win = what_am_i_looking_at(app)

        if screen != "edit_patient":
            status_callback("[PHASE 4] Clicking Add Pt...", "yellow")
            sel_win = app.top_window()

            added = False
            try:
                btns = sel_win.descendants(control_type="Button")
                for btn in btns:
                    txt = btn.window_text()
                    if "Add Pt" in txt:
                        btn.click_input()
                        added = True
                        break
            except Exception:
                pass

            if not added:
                pyautogui.hotkey('alt', 'a')

            time.sleep(3)

            # Handle popups
            screen, win = what_am_i_looking_at(app)
            while screen in ("trial_popup", "small_popup", "unknown_popup"):
                handle_screen(screen, win, patient, status_callback, app, None)
                time.sleep(1)
                try:
                    app = Application(backend="uia").connect(
                        title_re=".*Edit Patient.*|.*Open Dental.*",
                        timeout=5
                    )
                except Exception:
                    pass
                screen, win = what_am_i_looking_at(app)

        status_callback("[PHASE 4] Edit Patient form ready!", "limegreen")

        # ═══ PHASE 5: Fill the form ═══
        status_callback("[PHASE 5] Filling patient form...", "yellow")

        edit_win = app.top_window()
        edit_win.set_focus()
        time.sleep(0.5)

        fill_patient_form(edit_win, patient, status_callback)

        status_callback("[PHASE 5] Form filled!", "limegreen")

        # ═══ PHASE 6: Save ═══
        status_callback("[PHASE 6] Saving patient...", "yellow")

        saved = click_save(edit_win, status_callback)
        time.sleep(2)

        # Handle post-save popups
        screen, win = what_am_i_looking_at(app)
        if screen in ("trial_popup", "small_popup", "unknown_popup"):
            handle_screen(screen, win, patient, status_callback, app, None)

        if saved:
            status_callback(f"[DONE] {patient.first_name} {patient.last_name} saved in OpenDental!", "limegreen")
        else:
            status_callback("[DONE] Form filled but Save may not have clicked. Check OpenDental.", "orange")

        return saved

    except Exception as e:
        status_callback(f"ERROR: {e}", "red")
        return False
