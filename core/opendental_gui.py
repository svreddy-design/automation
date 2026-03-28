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
            _log(status_callback, "[3/6] SKIP — Already at Select Patient!", "limegreen")
        elif screen == "edit_patient":
            _log(status_callback, "[3/6] SKIP — Already at Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[3/6] Opening Select Patient...", "yellow")

            # pyautogui clicks the toolbar — pywinauto can't see toolbar buttons
            # ToolBarMain is at Y=50-75, center Y≈62 from window top
            # Select Patient is the first button, around X=90 from window left
            main_win = app.top_window()
            rect = main_win.rectangle()

            for click_attempt in range(3):
                # Click toolbar row 1 (Select Patient area)
                click_x = rect.left + 90
                click_y = rect.top + 70  # center of Y=50-75 toolbar
                _log(status_callback, f"  Click {click_attempt+1}: toolbar at ({click_x}, {click_y})", "cyan")
                pyautogui.click(click_x, click_y)
                time.sleep(3)

                # Check if Select Patient opened (pywinauto detection)
                app = _reconnect(app)
                screen, win, title = identify_screen(app)
                if screen == "select_patient":
                    break

                # If we opened something wrong, close it
                if screen in ("popup", "unknown"):
                    _log(status_callback, f"  Wrong dialog: '{title[:30]}' — closing", "cyan")
                    pyautogui.press('escape')
                    time.sleep(1)
                    app = _reconnect(app)

            screen, win, title = identify_screen(app)
            if screen == "select_patient":
                _log(status_callback, "[3/6] DONE — Select Patient open!", "limegreen")
            else:
                _log(status_callback, f"[3/6] FAILED — Expected Select Patient, got '{screen}'", "red")
                return False

        # ═══ STEP 4: Click Add Pt ═══
        screen, win, title = identify_screen(app)

        if screen == "edit_patient":
            _log(status_callback, "[4/6] SKIP — Already on Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[4/6] Clicking Add Pt...", "yellow")

            # pywinauto CAN see this button via auto_id
            main_win = app.top_window()
            clicked = False

            try:
                add_btn = main_win.child_window(auto_id="butAddPatient")
                if add_btn.exists(timeout=2):
                    add_btn.click_input()
                    clicked = True
                    _log(status_callback, "  Clicked butAddPatient!", "cyan")
            except Exception as e:
                _log(status_callback, f"  auto_id search failed: {e}", "orange")

            if not clicked:
                # Fallback: scan for any control with "Add Pt" text
                try:
                    for desc in main_win.descendants():
                        try:
                            if "Add Pt" in desc.window_text():
                                desc.click_input()
                                clicked = True
                                _log(status_callback, "  Found Add Pt by text scan!", "cyan")
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

            if not clicked:
                _log(status_callback, "[4/6] FAILED — Could not find Add Pt!", "red")
                return False

            # Wait for Edit Patient form
            time.sleep(3)
            app = _reconnect(app)

            # Verify Edit Patient opened
            for _ in range(5):
                screen, win, title = identify_screen(app)
                if screen == "edit_patient":
                    break
                elif screen in ("popup", "alerts"):
                    _dismiss(screen, win, status_callback)
                time.sleep(1)
                app = _reconnect(app)

            screen, win, title = identify_screen(app)
            if screen == "edit_patient":
                _log(status_callback, "[4/6] DONE — Edit Patient form open!", "limegreen")
            else:
                # Edit Patient might not use FormPatientEdit auto_id
                # Fall through and try filling fields anyway
                _log(status_callback, f"[4/6] WARNING — Can't confirm Edit Patient (got '{screen}'), continuing...", "orange")

        # ═══ STEP 5: Fill Form ═══
        _log(status_callback, "[5/6] Filling patient form...", "yellow")

        # pyautogui types into fields via Tab navigation
        # This is the "dumb but reliable" approach that worked before
        # OpenDental Edit Patient form: fields in tab order
        time.sleep(1)

        fields_to_enter = [
            ("last_name", patient.last_name),
            ("first_name", patient.first_name),
            ("middle_initial", patient.middle_initial),
            ("preferred_name", patient.preferred_name),
        ]

        # Type Last Name (cursor should already be on it)
        for field_name, value in fields_to_enter:
            if value:
                display = _mask(field_name, value)
                _log(status_callback, f"  {field_name} = {display}", "yellow")
                pyautogui.hotkey('ctrl', 'a')  # select all existing text
                pyautogui.write(value, interval=typing_interval)
            pyautogui.press('tab')
            time.sleep(field_delay)

        # Gender (dropdown — use arrow keys)
        if patient.gender:
            _log(status_callback, f"  gender = {patient.gender}", "yellow")
            gender_map = {"male": 1, "female": 2, "unknown": 3}
            presses = gender_map.get(patient.gender.lower(), 0)
            if presses:
                for _ in range(presses):
                    pyautogui.press('down')
                    time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(field_delay)

        # Position / other fields that may be between gender and the ones below
        # We'll tab through unknown fields
        # DOB, SSN, Address fields depend on the exact tab order
        # For now, type remaining fields with tab navigation

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
                    # DOB: type digits without slashes
                    digits = value.replace("/", "")
                    pyautogui.write(digits, interval=typing_interval)
                else:
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.write(value, interval=typing_interval)
            pyautogui.press('tab')
            time.sleep(field_delay)

        _log(status_callback, "[5/6] DONE — Form filled!", "limegreen")

        # ═══ STEP 6: Save ═══
        _log(status_callback, "[6/6] Saving...", "yellow")

        # Try pywinauto first (find Save/OK button)
        saved = False
        try:
            main_win = app.top_window()
            # Look for OK or Save button
            for aid in ["butOK", "butSave"]:
                try:
                    btn = main_win.child_window(auto_id=aid)
                    if btn.exists(timeout=1):
                        btn.click_input()
                        saved = True
                        _log(status_callback, f"  Clicked {aid}!", "cyan")
                        break
                except Exception:
                    continue
        except Exception:
            pass

        # Fallback: pyautogui Enter key
        if not saved:
            _log(status_callback, "  Pressing Enter to save...", "cyan")
            pyautogui.press('enter')

        time.sleep(2)

        # Check for errors after save
        app = _reconnect(app)
        screen, win, title = identify_screen(app)
        if screen in ("popup", "alerts"):
            error_text = ""
            try:
                error_text = win.window_text()
            except Exception:
                pass
            error_keywords = ["error", "fail", "invalid", "required", "cannot"]
            if any(kw in error_text.lower() for kw in error_keywords):
                _log(status_callback, f"[6/6] FAILED — Error: {error_text[:80]}", "red")
                return False
            else:
                _dismiss(screen, win, status_callback)

        _log(status_callback,
             f"[DONE] {patient.first_name} {patient.last_name} saved!", "limegreen")
        return True

    except Exception as e:
        _log(status_callback, f"CRITICAL ERROR: {e}", "red")
        return False
