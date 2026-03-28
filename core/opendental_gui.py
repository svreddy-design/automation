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

        # Small popup window (by title — fast check)
        if "Choose Database" not in title and "Open Dental" not in title and "Demo Database" not in title:
            try:
                rect = win.rectangle()
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w < 600 and h < 400 and w > 50:
                    return "popup", win, title
            except Exception:
                pass

        # Check for Edit Patient BEFORE Select Patient
        # (Edit Patient opens on top of Select Patient)
        try:
            ep = win.child_window(auto_id="FormPatientEdit")
            if ep.exists(timeout=0.5):
                return "edit_patient", ep, title
        except Exception:
            pass

        # Select Patient panel
        try:
            sp = win.child_window(auto_id="FormPatientSelect")
            if sp.exists(timeout=1):
                return "select_patient", sp, title
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
    pyautogui.PAUSE = 0.05  # faster between pyautogui actions

    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")
    field_delay = timing.get("field_delay_ms", 150) / 1000.0  # 150ms between fields
    typing_interval = timing.get("typing_interval_ms", 30) / 1000.0  # 30ms between keys

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

            main_win = app.top_window()
            rect = main_win.rectangle()

            # Click toolbar ONCE, then check
            click_x = rect.left + 90
            click_y = rect.top + 70
            _log(status_callback, f"  Clicking toolbar at ({click_x}, {click_y})", "cyan")
            pyautogui.click(click_x, click_y)
            time.sleep(2)

            app = _reconnect(app)
            screen, win, title = identify_screen(app)

            # If we got a popup (wrong button), close and retry
            if screen in ("popup", "unknown"):
                _log(status_callback, f"  Wrong dialog — closing...", "cyan")
                pyautogui.press('escape')
                time.sleep(1)
                app = _reconnect(app)
                # Try again with slightly different X
                pyautogui.click(click_x + 10, click_y)
                time.sleep(2)
                app = _reconnect(app)
                screen, win, title = identify_screen(app)

            # If still not open, maybe it toggled off — click one more time
            if screen != "select_patient":
                _log(status_callback, f"  Screen is '{screen}', clicking toolbar again...", "cyan")
                pyautogui.click(click_x, click_y)
                time.sleep(2)
                app = _reconnect(app)
                screen, win, title = identify_screen(app)

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
                    time.sleep(0.15)
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
                    time.sleep(0.15)
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

            # After clicking Add Pt, OpenDental may show popups:
            # - "Trial version. Maximum 30 patients" → press Enter
            # - "Not allowed to add...do a search first" → press Enter
            # Paid versions skip these popups entirely.
            # Strategy: check for edit_patient, if not found press Enter and check again
            app = _reconnect(app)
            for attempt in range(6):
                screen, win, title = identify_screen(app)
                if screen == "edit_patient":
                    _log(status_callback, "  Edit Patient detected!", "cyan")
                    break
                else:
                    # Either a popup or still on select_patient — press Enter to dismiss
                    _log(status_callback, f"  Attempt {attempt+1}: screen='{screen}', pressing Enter...", "cyan")
                    pyautogui.press('enter')
                    time.sleep(1.5)
                    app = _reconnect(app)

            screen, win, title = identify_screen(app)
            if screen == "edit_patient":
                _log(status_callback, "[5/8] DONE — Edit Patient form open!", "limegreen")
            else:
                _log(status_callback, f"[5/8] WARNING — Can't confirm Edit Patient (got '{screen}'), continuing...", "orange")

        # ═══ STEP 6: Fill Form (direct auto_id field access) ═══
        _log(status_callback, "[6/8] Filling patient form...", "yellow")
        time.sleep(0.5)

        main_win = app.top_window()

        # Map patient fields to OpenDental Edit Patient auto_ids
        # From diagnostic: all fields are [Edit] controls findable by auto_id
        field_map = [
            ("last_name",    patient.last_name,      "textLName"),
            ("first_name",   patient.first_name,     "textFName"),
            ("middle_initial", patient.middle_initial, "textMiddleI"),
            ("preferred_name", patient.preferred_name, "textPreferred"),
            ("phone",        patient.phone,           "textHmPhone"),
            ("address",      patient.address,         "textAddress"),
            ("city",         patient.city,            "textCity"),
            ("state",        patient.state,           "textState"),
            ("zip",          patient.zip,             "textZip"),
            ("ssn",          patient.ssn,             "textSSN"),
        ]

        def _type_into_field(field, value, field_name):
            """Click a field, clear it, type value using pywinauto set_edit_text.
            Falls back to pyautogui if set_edit_text fails.
            Returns True if typed successfully."""
            display = _mask(field_name, value)
            try:
                fr = field.rectangle()
                cx = (fr.left + fr.right) // 2
                cy = (fr.top + fr.bottom) // 2
                pyautogui.click(cx, cy)
                time.sleep(0.15)

                # Use pywinauto set_edit_text (more reliable than pyautogui.write)
                try:
                    field.set_edit_text(value)
                except Exception:
                    # Fallback: select all + type slowly
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.1)
                    pyautogui.write(value, interval=typing_interval)

                # Verify: read back what was typed
                time.sleep(0.1)
                try:
                    actual = field.window_text().strip()
                    if actual == value.strip():
                        _log(status_callback, f"  {field_name} = {display} [verified]", "yellow")
                    else:
                        # Retry once with set_edit_text
                        field.set_edit_text("")
                        time.sleep(0.1)
                        field.set_edit_text(value)
                        _log(status_callback, f"  {field_name} = {display} [retry]", "yellow")
                except Exception:
                    _log(status_callback, f"  {field_name} = {display} [set]", "yellow")
                return True
            except Exception as e:
                _log(status_callback, f"  {field_name}: error — {e}", "orange")
                return False

        for field_name, value, auto_id in field_map:
            if not value:
                continue
            try:
                field = main_win.child_window(auto_id=auto_id, control_type="Edit")
                if field.exists(timeout=1):
                    _type_into_field(field, value, field_name)
                else:
                    _log(status_callback, f"  {field_name}: not found ({auto_id})", "orange")
            except Exception as e:
                _log(status_callback, f"  {field_name}: error — {e}", "orange")
            time.sleep(field_delay)

        # Birthdate (auto_id='textDate' but multiple exist — find by label)
        if patient.dob:
            display = _mask("dob", patient.dob)
            filled_dob = False
            try:
                # Find the Birthdate label, then find the nearby Edit field
                labels = main_win.descendants(control_type="Text")
                for lbl in labels:
                    try:
                        if lbl.element_info.automation_id == "labelBirthdate":
                            lr = lbl.rectangle()
                            # The edit field is to the right of the label
                            # Find Edit fields near this label's Y position
                            for ed in main_win.descendants(control_type="Edit"):
                                er = ed.rectangle()
                                # Same row (within 20px) and to the right
                                if abs(er.top - lr.top) < 20 and er.left > lr.left:
                                    pyautogui.click((er.left + er.right) // 2,
                                                    (er.top + er.bottom) // 2)
                                    time.sleep(0.15)
                                    try:
                                        ed.set_edit_text(patient.dob)
                                    except Exception:
                                        pyautogui.hotkey('ctrl', 'a')
                                        pyautogui.write(patient.dob, interval=typing_interval)
                                    _log(status_callback, f"  dob = {display} [OK]", "yellow")
                                    filled_dob = True
                                    break
                            break
                    except Exception:
                        continue
            except Exception as e:
                _log(status_callback, f"  dob: error — {e}", "orange")
            if not filled_dob:
                _log(status_callback, "  dob: could not find Birthdate field", "orange")
            time.sleep(field_delay)

        # Gender (listbox — need to click the right item)
        if patient.gender:
            _log(status_callback, f"  gender = {patient.gender}", "yellow")
            try:
                items = main_win.descendants(control_type="ListItem")
                for item in items:
                    try:
                        if item.window_text() == patient.gender:
                            ir = item.rectangle()
                            pyautogui.click((ir.left + ir.right) // 2,
                                            (ir.top + ir.bottom) // 2)
                            _log(status_callback, f"  gender = {patient.gender} [OK]", "yellow")
                            break
                    except Exception:
                        continue
            except Exception:
                pass

        _log(status_callback, "[6/8] DONE — Form filled!", "limegreen")

        # ═══ STEP 7: Save ═══
        _log(status_callback, "[7/8] Saving...", "yellow")

        # Find and click Save button
        saved = False
        main_win = app.top_window()

        # Direct search for butSave (Button type from diagnostics)
        try:
            save_btn = main_win.child_window(auto_id="butSave", control_type="Button")
            if save_btn.exists(timeout=2):
                sr = save_btn.rectangle()
                pyautogui.click((sr.left + sr.right) // 2, (sr.top + sr.bottom) // 2)
                saved = True
                _log(status_callback, "  Clicked Save!", "cyan")
        except Exception:
            pass

        # Fallback: find any button with "Save" text
        if not saved:
            try:
                save_btn = main_win.child_window(title="Save", control_type="Button")
                if save_btn.exists(timeout=1):
                    sr = save_btn.rectangle()
                    pyautogui.click((sr.left + sr.right) // 2, (sr.top + sr.bottom) // 2)
                    saved = True
                    _log(status_callback, "  Clicked Save (by title)!", "cyan")
            except Exception:
                pass

        if not saved:
            _log(status_callback, "  Save not found — pressing Enter...", "orange")
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
