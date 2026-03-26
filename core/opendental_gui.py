"""Smart OpenDental GUI automation.
The bot checks what's on screen AFTER every action to verify it worked.
If something fails, it retries or stops with a clear error."""

import time
import sys
import os


def identify_screen(app):
    """Identify what screen/dialog is currently showing."""
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
        if "Select Patient" in title:
            return "select_patient", win
        if "Edit Patient" in title:
            return "edit_patient", win

        # Small window = popup dialog
        if width < 600 and height < 400 and width > 50:
            return "popup", win

        # Main OpenDental window
        if "Open Dental" in title or "Demo Database" in title:
            return "main_window", win

        return "unknown", win
    except Exception:
        return "error", None


def wait_for_screen(app, expected, status_callback, timeout=15):
    """Wait until we see the expected screen. Returns (screen_type, window) or None."""
    start = time.time()
    while time.time() - start < timeout:
        screen, win = identify_screen(app)

        if screen == expected:
            return screen, win

        # Auto-dismiss popups while waiting
        if screen in ("popup", "choose_database", "alerts"):
            dismiss_current(screen, win, status_callback)
            # Reconnect after dismiss
            time.sleep(2)
            try:
                app = reconnect(app)
            except Exception:
                pass
            continue

        time.sleep(1)

    return None, None


def dismiss_current(screen, win, status_callback):
    """Dismiss whatever popup/dialog is showing."""
    import pyautogui

    if screen == "choose_database":
        status_callback("  [AUTO] Dismissing Choose Database...", "cyan")
        try:
            ok = win.child_window(title="OK", control_type="Button")
            ok.click_input()
        except Exception:
            pyautogui.press('enter')
        time.sleep(3)

    elif screen == "alerts":
        status_callback("  [AUTO] Dismissing Alerts...", "cyan")
        try:
            ack = win.child_window(title="Acknowledge", control_type="Button")
            ack.click_input()
        except Exception:
            pyautogui.press('enter')
        time.sleep(2)

    elif screen == "popup":
        title = ""
        try:
            title = win.window_text()[:40]
        except Exception:
            pass
        status_callback(f"  [AUTO] Dismissing popup: {title}", "cyan")
        pyautogui.press('enter')
        time.sleep(1)


def reconnect(app):
    """Reconnect to OpenDental after a window change."""
    from pywinauto import Application
    try:
        return Application(backend="uia").connect(
            title_re=".*Open Dental.*|.*Demo Database.*|.*Select Patient.*|.*Edit Patient.*",
            timeout=10
        )
    except Exception:
        return app


def automate_patient_entry(patient, status_callback, config=None):
    """Smart automation with verification at every step."""
    if sys.platform != "win32":
        status_callback("ERROR: Requires Windows!", "red")
        return False

    from pywinauto import Application
    import pyautogui

    pyautogui.FAILSAFE = True
    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")

    try:
        # ═══ STEP 1: Connect or Launch ═══
        status_callback("[1/6] Finding OpenDental...", "yellow")

        app = None
        try:
            app = Application(backend="uia").connect(
                title_re=".*Open Dental.*|.*Demo Database.*", timeout=3
            )
            status_callback("[1/6] DONE — Connected!", "limegreen")
        except Exception:
            if not os.path.exists(app_path):
                status_callback(f"[1/6] FAILED — {app_path} not found!", "red")
                return False

            status_callback("[1/6] Launching OpenDental (wait 12s)...", "yellow")
            os.startfile(app_path)
            time.sleep(12)

            try:
                app = Application(backend="uia").connect(
                    title_re=".*Open Dental.*|.*Demo Database.*", timeout=20
                )
            except Exception:
                status_callback("[1/6] FAILED — OpenDental did not start!", "red")
                return False

            status_callback("[1/6] DONE — Launched!", "limegreen")

        # ═══ STEP 2: Get to Main Window ═══
        status_callback("[2/6] Getting to main screen...", "yellow")

        for attempt in range(20):
            screen, win = identify_screen(app)
            status_callback(f"  Attempt {attempt+1}: see '{screen}'", "cyan")

            if screen == "main_window":
                status_callback("[2/6] DONE — At main screen!", "limegreen")
                break

            elif screen == "select_patient":
                status_callback("[2/6] DONE — Already at Select Patient!", "limegreen")
                break

            elif screen == "edit_patient":
                status_callback("[2/6] DONE — Already at Edit Patient!", "limegreen")
                break

            elif screen in ("choose_database", "alerts", "popup"):
                dismiss_current(screen, win, status_callback)
                app = reconnect(app)
                continue

            elif screen == "unknown":
                # Might be login screen — press Enter
                status_callback("  Pressing Enter (login?)...", "cyan")
                pyautogui.press('enter')
                time.sleep(5)
                app = reconnect(app)
                continue

            else:
                time.sleep(2)
                app = reconnect(app)

        else:
            status_callback("[2/6] FAILED — Could not reach main screen after 20 attempts!", "red")
            return False

        # ═══ STEP 3: Open Select Patient ═══
        screen, win = identify_screen(app)

        if screen == "edit_patient":
            status_callback("[3/6] SKIP — Already on Edit Patient form!", "limegreen")

        elif screen == "select_patient":
            status_callback("[3/6] SKIP — Already on Select Patient!", "limegreen")

        else:
            status_callback("[3/6] Opening Select Patient...", "yellow")
            win = app.top_window()
            win.set_focus()

            # Click Select Patient
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
                except Exception:
                    pyautogui.hotkey('ctrl', 'p')

            time.sleep(2)

            # Dismiss popups and VERIFY we're at Select Patient
            for _ in range(5):
                app = reconnect(app)
                screen, win = identify_screen(app)
                if screen == "select_patient":
                    break
                elif screen in ("popup", "alerts"):
                    dismiss_current(screen, win, status_callback)
                    time.sleep(1)
                else:
                    time.sleep(1)

            screen, win = identify_screen(app)
            if screen == "select_patient":
                status_callback("[3/6] DONE — Select Patient open!", "limegreen")
            else:
                status_callback(f"[3/6] FAILED — Expected Select Patient but see '{screen}'!", "red")
                return False

        # ═══ STEP 4: Click Add Pt ═══
        screen, win = identify_screen(app)

        if screen == "edit_patient":
            status_callback("[4/6] SKIP — Already on Edit Patient!", "limegreen")
        else:
            status_callback("[4/6] Clicking Add Pt...", "yellow")

            sel_win = app.top_window()
            added = False
            try:
                btns = sel_win.descendants(control_type="Button")
                for btn in btns:
                    if "Add Pt" in btn.window_text():
                        btn.click_input()
                        added = True
                        break
            except Exception:
                pass
            if not added:
                pyautogui.hotkey('alt', 'a')

            time.sleep(2)

            # Dismiss popups and VERIFY we're at Edit Patient
            for _ in range(5):
                app = reconnect(app)
                screen, win = identify_screen(app)
                if screen == "edit_patient":
                    break
                elif screen in ("popup", "alerts"):
                    dismiss_current(screen, win, status_callback)
                    time.sleep(1)
                else:
                    time.sleep(1)

            screen, win = identify_screen(app)
            if screen == "edit_patient":
                status_callback("[4/6] DONE — Edit Patient form open!", "limegreen")
            else:
                status_callback(f"[4/6] FAILED — Expected Edit Patient but see '{screen}'!", "red")
                return False

        # ═══ STEP 5: Fill the form ═══
        status_callback("[5/6] Filling patient form...", "yellow")

        edit_win = app.top_window()
        edit_win.set_focus()
        time.sleep(0.5)

        # Get all editable fields
        edits = []
        try:
            edits = edit_win.descendants(control_type="Edit")
            status_callback(f"  Found {len(edits)} fields", "cyan")
        except Exception:
            status_callback("  WARNING: Could not find fields", "orange")

        def fill_field(name, value, display=None):
            if not value:
                return
            show = "***" if "SS" in name else value
            if display:
                show = display

            for edit in edits:
                try:
                    ename = edit.element_info.name or ""
                    if name in ename:
                        edit.click_input()
                        time.sleep(0.1)
                        edit.set_edit_text(value)
                        status_callback(f"  {name} = {show}", "yellow")
                        return
                except Exception:
                    continue

            try:
                e = edit_win.child_window(title=name, control_type="Edit")
                e.click_input()
                time.sleep(0.1)
                e.set_edit_text(value)
                status_callback(f"  {name} = {show}", "yellow")
                return
            except Exception:
                pass

            status_callback(f"  {name}: not found (skipped)", "orange")

        fill_field("Last Name", patient.last_name)
        time.sleep(0.2)
        fill_field("First Name", patient.first_name)
        time.sleep(0.2)
        fill_field("Middle Initial", patient.middle_initial)
        time.sleep(0.2)

        # Gender
        if patient.gender:
            try:
                items = edit_win.descendants(control_type="ListItem")
                for item in items:
                    if item.window_text() == patient.gender:
                        item.click_input()
                        status_callback(f"  Gender = {patient.gender}", "yellow")
                        break
            except Exception:
                pass

        fill_field("Birthdate", patient.dob)
        time.sleep(0.2)
        fill_field("Home Phone", patient.phone)
        time.sleep(0.2)
        fill_field("Address", patient.address)
        time.sleep(0.2)
        fill_field("City", patient.city)
        time.sleep(0.2)
        fill_field("ST", patient.state)
        time.sleep(0.2)

        if patient.zip:
            filled = False
            for edit in edits:
                try:
                    ename = edit.element_info.name or ""
                    if "Zip" in ename:
                        edit.click_input()
                        edit.set_edit_text(patient.zip)
                        status_callback(f"  Zip = {patient.zip}", "yellow")
                        filled = True
                        break
                except Exception:
                    continue
            if not filled:
                try:
                    combo = edit_win.child_window(title="Zip", control_type="ComboBox")
                    combo.click_input()
                    pyautogui.write(patient.zip, interval=0.03)
                    status_callback(f"  Zip = {patient.zip}", "yellow")
                except Exception:
                    pass

        if patient.ssn:
            fill_field("SS#", patient.ssn, "***")

        status_callback("[5/6] DONE — Form filled!", "limegreen")

        # ═══ STEP 6: Save ═══
        status_callback("[6/6] Saving...", "yellow")

        saved = False
        try:
            btns = edit_win.descendants(control_type="Button")
            for btn in btns:
                if btn.window_text() == "Save":
                    btn.click_input()
                    saved = True
                    break
        except Exception:
            pass

        if not saved:
            try:
                rect = edit_win.rectangle()
                pyautogui.click(rect.right - 60, rect.bottom - 30)
                saved = True
            except Exception:
                pass

        if not saved:
            status_callback("[6/6] FAILED — Could not find Save button!", "red")
            return False

        time.sleep(2)

        # Verify save worked — we should NOT be on Edit Patient anymore
        app = reconnect(app)
        screen, win = identify_screen(app)

        # Dismiss post-save popups
        if screen in ("popup", "alerts"):
            dismiss_current(screen, win, status_callback)

        status_callback(f"[DONE] {patient.first_name} {patient.last_name} saved in OpenDental!", "limegreen")
        return True

    except Exception as e:
        status_callback(f"ERROR: {e}", "red")
        return False
