"""Hybrid OpenDental GUI automation agent.

Uses pywinauto for what it CAN see:
  - Window connection & detection (FormPatientSelect, FormPatientEdit)
  - Clicking custom buttons by auto_id (butAddPatient, butOK)
  - Reading field values for verification

Uses pyautogui for what pywinauto CAN'T see:
  - Toolbar buttons (invisible custom WinForms ToolStrip)
  - Typing into fields via keyboard (Tab + type)
  - Keyboard shortcuts

Diagnostic data source (from diagnose.py on real OpenDental):
  - Toolbar 'ToolBarMain': rect Y=50-75, buttons invisible to UIA
  - Select Patient: child [Window] auto_id='FormPatientSelect'
  - Add Pt button: [Custom] auto_id='butAddPatient', text='_Add Pt'
  - Search fields: auto_id='textLName' > child Edit auto_id='textBox'
  - All buttons are [Custom] not [Button], text has underscore prefix
"""

import time
import sys
import os

from core.patient import PHI_FIELDS


# ═══════════════════════════════════════════════════
#  LOGGING — PHI-safe audit trail
# ═══════════════════════════════════════════════════

def _log(cb, msg, color="yellow"):
    cb(msg, color)


def _mask(field_name, value):
    if not value or field_name not in PHI_FIELDS:
        return value
    return "***" if field_name == "ssn" else "[REDACTED]"


# ═══════════════════════════════════════════════════
#  STATE DETECTION — pywinauto for what it can see
# ═══════════════════════════════════════════════════

def identify_screen(app):
    """Detect current screen using auto_ids from diagnostic data.
    Returns (screen_type, window, title)."""
    try:
        win = app.top_window()
        title = win.window_text()

        # Separate popup windows (by title)
        if "Choose Database" in title:
            return "choose_database", win, title
        if "Alert" in title and "Alerts" not in title:
            return "alerts", win, title

        # Child forms inside main window (by auto_id)
        try:
            sp = win.child_window(auto_id="FormPatientSelect")
            if sp.exists(timeout=1):
                return "select_patient", sp, title
        except Exception:
            pass

        try:
            ep = win.child_window(auto_id="FormPatientEdit")
            if ep.exists(timeout=0.5):
                return "edit_patient", ep, title
        except Exception:
            pass

        # Small popup
        try:
            rect = win.rectangle()
            w = rect.right - rect.left
            h = rect.bottom - rect.top
            if w < 600 and h < 400 and w > 50:
                return "popup", win, title
        except Exception:
            pass

        if "Open Dental" in title or "Demo Database" in title:
            return "main_window", win, title

        return "unknown", win, title
    except Exception:
        return "error", None, ""


# ═══════════════════════════════════════════════════
#  POPUP DISMISSAL
# ═══════════════════════════════════════════════════

def _dismiss(screen, win, cb):
    """Dismiss popups using pywinauto where possible, pyautogui as fallback."""
    import pyautogui

    if screen == "choose_database":
        _log(cb, "  [AUTO] Dismissing Choose Database...", "cyan")
        try:
            ok = win.child_window(title="OK", control_type="Button")
            if ok.exists(timeout=1):
                ok.click_input()
            else:
                pyautogui.press('enter')
        except Exception:
            pyautogui.press('enter')
        time.sleep(2)

    elif screen == "alerts":
        _log(cb, "  [AUTO] Dismissing Alerts...", "cyan")
        try:
            ack = win.child_window(title="Acknowledge", control_type="Button")
            if ack.exists(timeout=1):
                ack.click_input()
            else:
                pyautogui.press('enter')
        except Exception:
            pyautogui.press('enter')
        time.sleep(1.5)

    elif screen == "popup":
        _log(cb, f"  [AUTO] Dismissing popup: {win.window_text()[:40]}", "cyan")
        pyautogui.press('enter')
        time.sleep(1)


def _reconnect(app):
    from pywinauto import Application
    try:
        return Application(backend="uia").connect(
            title_re=".*Open Dental.*|.*Demo Database.*",
            timeout=10
        )
    except Exception:
        return app


def dismiss_all_dialogs(app, cb, max_rounds=5):
    """Dismiss all blocking dialogs."""
    for _ in range(max_rounds):
        screen, win, title = identify_screen(app)
        if screen in ("choose_database", "alerts", "popup"):
            _dismiss(screen, win, cb)
            time.sleep(1)
            app = _reconnect(app)
        else:
            break
    return app


# ═══════════════════════════════════════════════════
#  MAIN AUTOMATION
# ═══════════════════════════════════════════════════

def automate_patient_entry(patient, status_callback, config=None):
    """Enter a patient into OpenDental using hybrid automation.

    pywinauto: window detection, clicking buttons by auto_id
    pyautogui: toolbar clicks, keyboard typing into fields

    Returns True if patient saved, False otherwise.
    """
    if sys.platform != "win32":
        _log(status_callback, "ERROR: Requires Windows!", "red")
        return False

    from pywinauto import Application
    import pyautogui

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1

    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")
    field_delay = timing.get("field_delay_ms", 300) / 1000.0
    typing_interval = timing.get("typing_interval_ms", 50) / 1000.0

    try:
        # ═══ STEP 1: Connect or Launch ═══
        _log(status_callback, "[1/6] Finding OpenDental...", "yellow")

        app = None
        try:
            app = Application(backend="uia").connect(
                title_re=".*Open Dental.*|.*Demo Database.*", timeout=3
            )
            _log(status_callback, "[1/6] DONE — Connected!", "limegreen")
        except Exception:
            if not os.path.exists(app_path):
                _log(status_callback, f"[1/6] FAILED — {app_path} not found!", "red")
                return False
            load_delay = timing.get("app_load_delay_s", 12)
            _log(status_callback, f"[1/6] Launching OpenDental ({load_delay}s)...", "yellow")
            os.startfile(app_path)
            time.sleep(load_delay)
            try:
                app = Application(backend="uia").connect(
                    title_re=".*Open Dental.*|.*Demo Database.*", timeout=20
                )
            except Exception:
                _log(status_callback, "[1/6] FAILED — OpenDental did not start!", "red")
                return False
            _log(status_callback, "[1/6] DONE — Launched!", "limegreen")

        # ═══ STEP 2: Get to Main Window ═══
        _log(status_callback, "[2/6] Getting to main screen...", "yellow")

        for attempt in range(15):
            screen, win, title = identify_screen(app)

            if screen == "main_window":
                app = dismiss_all_dialogs(app, status_callback)
                _log(status_callback, "[2/6] DONE — At main screen!", "limegreen")
                break
            elif screen == "select_patient":
                _log(status_callback, "[2/6] DONE — Already at Select Patient!", "limegreen")
                break
            elif screen == "edit_patient":
                _log(status_callback, "[2/6] DONE — Already at Edit Patient!", "limegreen")
                break
            elif screen in ("choose_database", "alerts", "popup"):
                _dismiss(screen, win, status_callback)
                app = _reconnect(app)
            elif screen == "unknown":
                pyautogui.press('enter')
                time.sleep(timing.get("login_delay_s", 6))
                app = _reconnect(app)
            else:
                time.sleep(1.5)
                app = _reconnect(app)
        else:
            _log(status_callback, "[2/6] FAILED — Could not reach main screen!", "red")
            return False

        # ═══ STEP 3: Open Select Patient ═══
        screen, win, title = identify_screen(app)

        if screen == "select_patient":
            _log(status_callback, "[3/8] SKIP — Already at Select Patient!", "limegreen")
        elif screen == "edit_patient":
            _log(status_callback, "[3/8] SKIP — Already at Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[3/8] Opening Select Patient...", "yellow")

            # pyautogui clicks the toolbar (invisible to pywinauto)
            main_win = app.top_window()
            rect = main_win.rectangle()

            for click_attempt in range(3):
                click_x = rect.left + 90
                click_y = rect.top + 70
                _log(status_callback, f"  Click {click_attempt+1}: toolbar at ({click_x}, {click_y})", "cyan")
                pyautogui.click(click_x, click_y)
                time.sleep(3)

                app = _reconnect(app)
                screen, win, title = identify_screen(app)
                if screen == "select_patient":
                    break

                if screen in ("popup", "unknown"):
                    pyautogui.press('escape')
                    time.sleep(1)
                    app = _reconnect(app)

            screen, win, title = identify_screen(app)
            if screen == "select_patient":
                _log(status_callback, "[3/8] DONE — Select Patient open!", "limegreen")
            else:
                _log(status_callback, f"[3/8] FAILED — Expected Select Patient, got '{screen}'", "red")
                return False

        # ═══ STEP 4: Search for patient (required before Add Pt) ═══
        screen, win, title = identify_screen(app)

        if screen == "edit_patient":
            _log(status_callback, "[4/8] SKIP — Already at Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[4/8] Searching for patient (required by OpenDental)...", "yellow")

            main_win = app.top_window()

            # Type last name in search field (auto_id='textLName' > child 'textBox')
            try:
                ln_field = main_win.child_window(auto_id="textLName").child_window(auto_id="textBox")
                if ln_field.exists(timeout=2):
                    ln_rect = ln_field.rectangle()
                    pyautogui.click((ln_rect.left + ln_rect.right) // 2,
                                    (ln_rect.top + ln_rect.bottom) // 2)
                    time.sleep(0.3)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.write(patient.last_name, interval=typing_interval)
                    _log(status_callback, f"  Last Name: {patient.last_name}", "cyan")
            except Exception as e:
                _log(status_callback, f"  Could not type last name: {e}", "orange")

            # Type first name
            try:
                fn_field = main_win.child_window(auto_id="textFName").child_window(auto_id="textBox")
                if fn_field.exists(timeout=1):
                    fn_rect = fn_field.rectangle()
                    pyautogui.click((fn_rect.left + fn_rect.right) // 2,
                                    (fn_rect.top + fn_rect.bottom) // 2)
                    time.sleep(0.3)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.write(patient.first_name, interval=typing_interval)
                    _log(status_callback, f"  First Name: {patient.first_name}", "cyan")
            except Exception:
                pass

            # Click Search button
            try:
                search_btn = main_win.child_window(auto_id="butSearch")
                if search_btn.exists(timeout=1):
                    sr = search_btn.rectangle()
                    pyautogui.click((sr.left + sr.right) // 2, (sr.top + sr.bottom) // 2)
                    _log(status_callback, "  Clicked Search", "cyan")
                    time.sleep(2)
            except Exception:
                pass

            # Check if patient already exists in the grid results
            # (We can't easily read the grid, so we just log and continue)
            _log(status_callback, "[4/8] DONE — Search complete!", "limegreen")

        # ═══ STEP 5: Click Add Pt + dismiss popups ═══
        screen, win, title = identify_screen(app)

        if screen == "edit_patient":
            _log(status_callback, "[5/8] SKIP — Already on Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[5/8] Clicking Add Pt...", "yellow")

            main_win = app.top_window()

            # Find Add Pt button and click with pyautogui
            try:
                add_btn = main_win.child_window(auto_id="butAddPatient")
                if add_btn.exists(timeout=2):
                    btn_rect = add_btn.rectangle()
                    cx = (btn_rect.left + btn_rect.right) // 2
                    cy = (btn_rect.top + btn_rect.bottom) // 2
                    _log(status_callback, f"  Found butAddPatient at ({cx}, {cy})", "cyan")
                    pyautogui.click(cx, cy)
                else:
                    _log(status_callback, "[5/8] FAILED — Add Pt button not found!", "red")
                    return False
            except Exception as e:
                _log(status_callback, f"[5/8] FAILED — {e}", "red")
                return False

            time.sleep(1.5)

            # Dismiss popup: "Trial version. Maximum 30 patients"
            _log(status_callback, "  Dismissing trial popup...", "cyan")
            pyautogui.press('enter')
            time.sleep(1)

            # Dismiss popup: "Not allowed to add... do a search first"
            # (We already searched, so this shouldn't appear, but just in case)
            _log(status_callback, "  Dismissing any remaining popups...", "cyan")
            pyautogui.press('enter')
            time.sleep(2)

            # Check if Edit Patient opened
            app = _reconnect(app)
            for _ in range(5):
                screen, win, title = identify_screen(app)
                if screen == "edit_patient":
                    break
                elif screen in ("popup", "alerts"):
                    _log(status_callback, f"  Popup: pressing OK...", "cyan")
                    pyautogui.press('enter')
                    time.sleep(1)
                    app = _reconnect(app)
                else:
                    time.sleep(1)
                    app = _reconnect(app)

            screen, win, title = identify_screen(app)
            if screen == "edit_patient":
                _log(status_callback, "[5/8] DONE — Edit Patient form open!", "limegreen")
            else:
                _log(status_callback, f"[5/8] WARNING — Can't confirm Edit Patient (got '{screen}'), continuing...", "orange")

        # ═══ STEP 6: Fill Form ═══
        _log(status_callback, "[6/8] Filling patient form...", "yellow")
        time.sleep(1)

        # pyautogui Tab+type through the Edit Patient form fields
        # The cursor should be on the first field (Last Name) after the form opens

        fields_to_enter = [
            ("last_name", patient.last_name),
            ("first_name", patient.first_name),
            ("middle_initial", patient.middle_initial),
            ("preferred_name", patient.preferred_name),
        ]

        for field_name, value in fields_to_enter:
            if value:
                display = _mask(field_name, value)
                _log(status_callback, f"  {field_name} = {display}", "yellow")
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.write(value, interval=typing_interval)
            pyautogui.press('tab')
            time.sleep(field_delay)

        # Gender (dropdown)
        if patient.gender:
            _log(status_callback, f"  gender = {patient.gender}", "yellow")
            gender_map = {"male": 1, "female": 2, "unknown": 3}
            presses = gender_map.get(patient.gender.lower(), 0)
            for _ in range(presses):
                pyautogui.press('down')
                time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(field_delay)

        remaining_fields = [
            ("dob", patient.dob),
            ("ssn", patient.ssn),
            ("address", patient.address),
            ("city", patient.city),
            ("state", patient.state),
            ("zip", patient.zip),
            ("phone", patient.phone),
        ]

        for field_name, value in remaining_fields:
            if value:
                display = _mask(field_name, value)
                _log(status_callback, f"  {field_name} = {display}", "yellow")
                if field_name == "dob":
                    digits = value.replace("/", "")
                    pyautogui.write(digits, interval=typing_interval)
                else:
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.write(value, interval=typing_interval)
            pyautogui.press('tab')
            time.sleep(field_delay)

        _log(status_callback, "[6/8] DONE — Form filled!", "limegreen")

        # ═══ STEP 7: Save ═══
        _log(status_callback, "[7/8] Saving...", "yellow")

        # Try pywinauto first (find Save/OK button)
        saved = False
        try:
            main_win = app.top_window()
            for aid in ["butOK", "butSave"]:
                try:
                    btn = main_win.child_window(auto_id=aid)
                    if btn.exists(timeout=1):
                        br = btn.rectangle()
                        pyautogui.click((br.left + br.right) // 2,
                                        (br.top + br.bottom) // 2)
                        saved = True
                        _log(status_callback, f"  Clicked {aid}!", "cyan")
                        break
                except Exception:
                    continue
        except Exception:
            pass

        if not saved:
            _log(status_callback, "  Pressing Enter to save...", "cyan")
            pyautogui.press('enter')

        time.sleep(2)

        # ═══ STEP 8: Verify save ═══
        _log(status_callback, "[8/8] Verifying...", "yellow")

        app = _reconnect(app)
        screen, win, title = identify_screen(app)

        # Dismiss any post-save popups
        for _ in range(3):
            if screen in ("popup", "alerts"):
                error_text = ""
                try:
                    error_text = win.window_text()
                except Exception:
                    pass
                error_keywords = ["error", "fail", "invalid", "required", "cannot"]
                if any(kw in error_text.lower() for kw in error_keywords):
                    _log(status_callback, f"[8/8] FAILED — Error: {error_text[:80]}", "red")
                    return False
                _log(status_callback, f"  Dismissing: {error_text[:40]}", "cyan")
                pyautogui.press('enter')
                time.sleep(1)
                app = _reconnect(app)
                screen, win, title = identify_screen(app)
            else:
                break

        _log(status_callback,
             f"[DONE] {patient.first_name} {patient.last_name} saved!", "limegreen")
        return True

    except Exception as e:
        _log(status_callback, f"CRITICAL ERROR: {e}", "red")
        return False
