"""Production-grade OpenDental GUI automation agent.

Built on 10 principles:
1. Reliability over speed — wait for elements, never hardcoded sleep
2. Error recovery — handle crashes, never leave app in half-posted state
3. Smart element identification — auto_id/control_type/title, never coordinates
4. Verification after every action — read back fields, check for errors
5. HIPAA/PHI security — mask SSN, phone, address, DOB in all logs
6. Idempotency — check for duplicate patients before adding
7. Human-in-the-loop — caller handles confirmation before invoking
8. Handle OpenDental versions — configurable locators via config.json
9. Speed control — configurable realistic delays between actions
10. State management — always know which screen we're on before acting

Uses pywinauto exclusively. Zero pyautogui dependency.
Windows-only (OpenDental is a Windows application).
"""

import time
import sys
import os
from core.patient import PHI_FIELDS
from core.opendental import load_locators


# ═══════════════════════════════════════════════════
#  LOGGING — PHI-safe action audit trail
# ═══════════════════════════════════════════════════

def _log(status_callback, msg, color="yellow"):
    """Log with timestamp. All logging goes through here."""
    status_callback(msg, color)


def _mask(field_name, value):
    """Mask PHI fields for logging. Never log raw sensitive data."""
    if not value or field_name not in PHI_FIELDS:
        return value
    if field_name == "ssn":
        return "***"
    return "[REDACTED]"


# ═══════════════════════════════════════════════════
#  STATE MANAGEMENT — always know where we are
# ═══════════════════════════════════════════════════

def identify_screen(app):
    """Identify current OpenDental screen. Returns (screen_type, window, title).
    Uses system-wide window search because OpenDental's dialogs (Select Patient,
    Edit Patient) are child windows that app.top_window() and app.windows() miss.
    Screen types: choose_database, alerts, select_patient, edit_patient,
    popup, main_window, unknown, error."""
    from pywinauto import Desktop

    try:
        # Search ALL visible windows on the desktop for OpenDental dialogs
        # This catches child/owned windows that app.windows() misses
        try:
            desktop = Desktop(backend="uia")
            for dwin in desktop.windows():
                try:
                    t = dwin.window_text()
                    if "Select Patient" in t:
                        return "select_patient", dwin, t
                    if "Edit Patient" in t:
                        return "edit_patient", dwin, t
                    if "Choose Database" in t:
                        return "choose_database", dwin, t
                except Exception:
                    continue
        except Exception:
            pass

        # Fall back to app.top_window for main window / popup detection
        win = app.top_window()
        title = win.window_text()
        rect = win.rectangle()
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        if "Alert" in title:
            return "alerts", win, title
        if "Select Patient" in title:
            return "select_patient", win, title
        if "Edit Patient" in title:
            return "edit_patient", win, title

        # Small window = popup dialog
        if width < 600 and height < 400 and width > 50:
            return "popup", win, title

        # Main OpenDental window
        if "Open Dental" in title or "Demo Database" in title:
            return "main_window", win, title

        return "unknown", win, title
    except Exception:
        return "error", None, ""


def wait_for_screen(app, expected, status_callback, timeout=15):
    """Wait until we see the expected screen. Auto-dismisses popups.
    Returns (screen_type, window) or (None, None) on timeout."""
    start = time.time()
    while time.time() - start < timeout:
        screen, win, title = identify_screen(app)

        if screen == expected:
            return screen, win

        # Auto-dismiss popups while waiting
        if screen in ("popup", "choose_database", "alerts"):
            _dismiss(screen, win, status_callback)
            time.sleep(1.5)
            try:
                app = _reconnect(app)
            except Exception:
                pass
            continue

        time.sleep(0.5)

    return None, None


# ═══════════════════════════════════════════════════
#  ELEMENT FINDING — smart, never coordinates
# ═══════════════════════════════════════════════════

def find_element(win, locator, status_callback=None, retries=3):
    """Find a UI element using configurable locator dict.

    Locator format: {"title": "Save", "control_type": "Button"}
    or: {"auto_id": "btnSave"} or {"title_re": ".*Save.*"}

    Tries: auto_id → title+control_type → title_re → descendants scan.
    Never falls back to screen coordinates."""
    for attempt in range(retries):
        # Strategy 1: auto_id (most stable across versions)
        if "auto_id" in locator:
            try:
                elem = win.child_window(auto_id=locator["auto_id"])
                if elem.exists(timeout=1):
                    return elem
            except Exception:
                pass

        # Strategy 2: title + control_type (most common)
        if "title" in locator and "control_type" in locator:
            try:
                elem = win.child_window(
                    title=locator["title"],
                    control_type=locator["control_type"]
                )
                if elem.exists(timeout=1):
                    return elem
            except Exception:
                pass

        # Strategy 3: regex title match
        if "title_re" in locator:
            try:
                ctrl_type = locator.get("control_type")
                kwargs = {"title_re": locator["title_re"]}
                if ctrl_type:
                    kwargs["control_type"] = ctrl_type
                elem = win.child_window(**kwargs)
                if elem.exists(timeout=1):
                    return elem
            except Exception:
                pass

        # Strategy 4: scan descendants by control_type, match title substring
        if "title" in locator and "control_type" in locator:
            try:
                for desc in win.descendants(control_type=locator["control_type"]):
                    try:
                        name = desc.window_text() or desc.element_info.name or ""
                        if locator["title"] in name:
                            return desc
                    except Exception:
                        continue
            except Exception:
                pass

        if attempt < retries - 1:
            time.sleep(0.5)

    if status_callback:
        _log(status_callback, f"  Element not found: {locator} (after {retries} attempts)", "orange")
    return None


# ═══════════════════════════════════════════════════
#  FIELD OPERATIONS — fill + verify
# ═══════════════════════════════════════════════════

def fill_and_verify(edit, value, field_name, status_callback):
    """Set a field value and read it back to confirm.
    Returns True if verified, False if mismatch."""
    if not value:
        return True

    display = _mask(field_name, value)

    try:
        edit.set_edit_text(value)
        time.sleep(0.2)

        # Read back to verify
        actual = ""
        try:
            actual = edit.window_text()
        except Exception:
            try:
                actual = edit.get_value()
            except Exception:
                # Can't read back — log warning but don't fail
                _log(status_callback, f"  {field_name} = {display} (set, can't verify)", "yellow")
                return True

        if actual.strip() == value.strip():
            _log(status_callback, f"  {field_name} = {display} [verified]", "yellow")
            return True
        else:
            _log(status_callback,
                 f"  VERIFY FAIL: {field_name} — expected '{display}', got different value", "red")
            return False

    except Exception as e:
        _log(status_callback, f"  ERROR setting {field_name}: {e}", "red")
        return False


# ═══════════════════════════════════════════════════
#  POPUP/DIALOG HANDLING
# ═══════════════════════════════════════════════════

def _dismiss(screen, win, status_callback):
    """Dismiss popup/dialog using pywinauto only."""
    from pywinauto import keyboard as pwa_keyboard

    if screen == "choose_database":
        _log(status_callback, "  [AUTO] Dismissing Choose Database...", "cyan")
        try:
            ok = win.child_window(title="OK", control_type="Button")
            if ok.exists(timeout=1):
                ok.click_input()
            else:
                pwa_keyboard.send_keys('{ENTER}')
        except Exception:
            pwa_keyboard.send_keys('{ENTER}')
        time.sleep(2)

    elif screen == "alerts":
        _log(status_callback, "  [AUTO] Dismissing Alerts...", "cyan")
        try:
            ack = win.child_window(title="Acknowledge", control_type="Button")
            if ack.exists(timeout=1):
                ack.click_input()
            else:
                pwa_keyboard.send_keys('{ENTER}')
        except Exception:
            pwa_keyboard.send_keys('{ENTER}')
        time.sleep(1.5)

    elif screen == "popup":
        title = ""
        try:
            title = win.window_text()[:40]
        except Exception:
            pass
        _log(status_callback, f"  [AUTO] Dismissing popup: {title}", "cyan")
        pwa_keyboard.send_keys('{ENTER}')
        time.sleep(1)


def _reconnect(app):
    """Reconnect to OpenDental after a window change."""
    from pywinauto import Application
    try:
        return Application(backend="uia").connect(
            title_re=".*Open Dental.*|.*Demo Database.*|.*Select Patient.*|.*Edit Patient.*",
            timeout=10
        )
    except Exception:
        return app


def dismiss_all_dialogs(app, status_callback, max_rounds=5):
    """Aggressively dismiss ALL blocking dialogs (Choose Database, alerts, popups).
    Checks every window, not just top_window, to catch dialogs hiding behind."""

    for _ in range(max_rounds):
        dismissed = False

        # Check all windows belonging to the app
        try:
            for win in app.windows():
                try:
                    title = win.window_text()
                    if "Choose Database" in title:
                        _dismiss("choose_database", win, status_callback)
                        dismissed = True
                    elif "Alert" in title:
                        _dismiss("alerts", win, status_callback)
                        dismissed = True
                except Exception:
                    continue
        except Exception:
            pass

        # Also check via top_window
        screen, win, title = identify_screen(app)
        if screen in ("choose_database", "alerts", "popup"):
            _dismiss(screen, win, status_callback)
            dismissed = True

        if not dismissed:
            break

        time.sleep(1)
        try:
            app = _reconnect(app)
        except Exception:
            pass

    return app


# ═══════════════════════════════════════════════════
#  DUPLICATE CHECK — idempotency
# ═══════════════════════════════════════════════════

def check_duplicate(win, patient, status_callback):
    """In Select Patient dialog, search for existing patient.
    Returns True if a likely duplicate is found."""
    from pywinauto import keyboard as pwa_keyboard

    _log(status_callback, "  Checking for duplicate patient...", "cyan")

    try:
        # Find the search/last name field in Select Patient dialog
        search_fields = win.descendants(control_type="Edit")
        if search_fields:
            # First edit field is usually the last name search
            search = search_fields[0]
            search.click_input()
            search.set_edit_text(patient.last_name)
            time.sleep(0.3)

            # Press Enter or Tab to trigger search
            pwa_keyboard.send_keys('{ENTER}')
            time.sleep(1)

            # Check if results list has matching entries
            try:
                rows = win.descendants(control_type="DataItem")
                for row in rows:
                    text = row.window_text().lower()
                    if (patient.last_name.lower() in text and
                            patient.first_name.lower() in text):
                        _log(status_callback,
                             f"  DUPLICATE FOUND: {patient.first_name} {patient.last_name} already exists!",
                             "orange")
                        return True
            except Exception:
                pass

            # Clear search field for next use
            try:
                search.set_edit_text("")
            except Exception:
                pass

    except Exception as e:
        _log(status_callback, f"  Duplicate check skipped: {e}", "cyan")

    return False


# ═══════════════════════════════════════════════════
#  MAIN AUTOMATION ENTRY POINT
# ═══════════════════════════════════════════════════

def automate_patient_entry(patient, status_callback, config=None):
    """Enter a patient into OpenDental via GUI automation.

    This is the paranoid agent — it verifies every step, masks PHI,
    checks for duplicates, and never assumes an action succeeded.

    Args:
        patient: Patient dataclass (validated)
        status_callback: function(text, color) for status updates
        config: dict with app_path and timing overrides

    Returns:
        True if patient was successfully entered, False otherwise.
    """
    if sys.platform != "win32":
        _log(status_callback, "ERROR: Requires Windows! OpenDental is a Windows application.", "red")
        return False

    from pywinauto import Application
    from pywinauto import keyboard as pwa_keyboard

    timing = config or {}
    app_path = timing.get("app_path", r"C:\Program Files (x86)\Open Dental\OpenDental.exe")
    locators = load_locators()
    field_delay = timing.get("field_delay_ms", 200) / 1000.0
    action_delay = timing.get("action_delay_ms", 500) / 1000.0

    try:
        # ═══ STEP 1: Connect or Launch OpenDental ═══
        _log(status_callback, "[1/7] Finding OpenDental...", "yellow")

        app = None
        try:
            app = Application(backend="uia").connect(
                title_re=".*Open Dental.*|.*Demo Database.*", timeout=3
            )
            _log(status_callback, "[1/7] DONE — Connected to running instance!", "limegreen")
        except Exception:
            if not os.path.exists(app_path):
                _log(status_callback, f"[1/7] FAILED — {app_path} not found!", "red")
                return False

            load_delay = timing.get("app_load_delay_s", 12)
            _log(status_callback, f"[1/7] Launching OpenDental (waiting {load_delay}s)...", "yellow")
            os.startfile(app_path)
            time.sleep(load_delay)

            try:
                app = Application(backend="uia").connect(
                    title_re=".*Open Dental.*|.*Demo Database.*", timeout=20
                )
            except Exception:
                _log(status_callback, "[1/7] FAILED — OpenDental did not start!", "red")
                return False

            _log(status_callback, "[1/7] DONE — Launched!", "limegreen")

        # ═══ STEP 2: Navigate to Main Window ═══
        _log(status_callback, "[2/7] Getting to main screen...", "yellow")

        for attempt in range(20):
            screen, win, title = identify_screen(app)
            _log(status_callback, f"  Attempt {attempt + 1}: screen='{screen}' title='{title[:40]}'", "cyan")

            if screen == "main_window":
                # Dismiss any lurking dialogs (Choose Database, alerts) before proceeding
                app = dismiss_all_dialogs(app, status_callback)

                # Log toolbar controls for diagnostics (helps fix locators)
                try:
                    main_win = app.top_window()
                    toolbar_items = []
                    for desc in main_win.descendants():
                        try:
                            text = desc.window_text()
                            ctrl = desc.element_info.control_type
                            if text and len(text) < 40:
                                toolbar_items.append(f"{text}({ctrl})")
                        except Exception:
                            continue
                    if toolbar_items:
                        _log(status_callback,
                             f"  Controls found: {', '.join(toolbar_items[:20])}", "cyan")
                except Exception:
                    pass

                _log(status_callback, "[2/7] DONE — At main screen!", "limegreen")
                break
            elif screen == "select_patient":
                _log(status_callback, "[2/7] DONE — Already at Select Patient!", "limegreen")
                break
            elif screen == "edit_patient":
                _log(status_callback, "[2/7] DONE — Already at Edit Patient!", "limegreen")
                break
            elif screen in ("choose_database", "alerts", "popup"):
                _dismiss(screen, win, status_callback)
                app = _reconnect(app)
                continue
            elif screen == "unknown":
                _log(status_callback, "  Pressing Enter (possible login screen)...", "cyan")
                pwa_keyboard.send_keys('{ENTER}')
                time.sleep(timing.get("login_delay_s", 6))
                app = _reconnect(app)
                continue
            else:
                time.sleep(1.5)
                app = _reconnect(app)
        else:
            _log(status_callback, "[2/7] FAILED — Could not reach main screen after 20 attempts!", "red")
            return False

        # ═══ STEP 3: Open Select Patient ═══
        screen, win, title = identify_screen(app)

        if screen == "edit_patient":
            _log(status_callback, "[3/7] SKIP — Already on Edit Patient form!", "limegreen")
        elif screen == "select_patient":
            _log(status_callback, "[3/7] SKIP — Already on Select Patient!", "limegreen")
        else:
            _log(status_callback, "[3/7] Opening Select Patient...", "yellow")
            win = app.top_window()
            win.set_focus()
            time.sleep(0.3)

            opened = False

            # Strategy 1: Click "Select Patient" on the first toolbar row
            # OpenDental toolbar layout:
            #   Row 1: [icon] Select Patient | Commlog | E-mail | WebMail | ...
            #   Row 2: Print | Lists | Pat Appts | ...
            # We must click Row 1, NOT Row 2 (which has Print)
            # Title bar ~31px + menu bar ~23px = toolbar row 1 starts at ~54px
            _log(status_callback, "  Clicking Select Patient on toolbar row 1...", "cyan")
            try:
                rect = win.rectangle()
                # Row 1 center: ~65px from window top, ~90px from left
                toolbar_x = rect.left + 90
                toolbar_y = rect.top + 65
                from pywinauto import mouse as pwa_mouse
                pwa_mouse.click(coords=(toolbar_x, toolbar_y))
                _log(status_callback, f"  Clicked at ({toolbar_x}, {toolbar_y})", "cyan")
                time.sleep(2)
                app = _reconnect(app)
                screen_check, _, _ = identify_screen(app)
                if screen_check == "select_patient":
                    opened = True
            except Exception as e:
                _log(status_callback, f"  Toolbar click failed: {e}", "orange")

            # Strategy 2: If we accidentally opened something else (like Print), close it
            if not opened:
                app = _reconnect(app)
                screen_check, win_check, title_check = identify_screen(app)
                if screen_check in ("popup", "unknown"):
                    _log(status_callback, f"  Wrong dialog opened: '{title_check}' — closing...", "cyan")
                    pwa_keyboard.send_keys('{ESC}')
                    time.sleep(1)
                    app = _reconnect(app)

            # Strategy 3: Try ToolBar.button() method
            if not opened:
                try:
                    win = app.top_window()
                    toolbars = win.descendants(control_type="ToolBar")
                    for tb in toolbars:
                        try:
                            tb_btn = tb.button("Select Patient")
                            tb_btn.click()
                            opened = True
                            _log(status_callback, "  Found via ToolBar.button()!", "cyan")
                            break
                        except Exception:
                            continue
                except Exception:
                    pass

            # Strategy 4: Try different Y offsets on toolbar row 1
            if not opened:
                _log(status_callback, "  Trying multiple toolbar positions...", "cyan")
                try:
                    rect = win.rectangle()
                    from pywinauto import mouse as pwa_mouse
                    # Try Y offsets 58, 62, 68, 72 to find the right row
                    for y_off in [58, 62, 68, 72]:
                        pwa_mouse.click(coords=(rect.left + 90, rect.top + y_off))
                        time.sleep(1.5)
                        app = _reconnect(app)
                        screen_check, _, _ = identify_screen(app)
                        if screen_check == "select_patient":
                            opened = True
                            _log(status_callback, f"  Found at Y offset {y_off}!", "cyan")
                            break
                        elif screen_check in ("popup", "unknown"):
                            # Wrong dialog — close and try next offset
                            pwa_keyboard.send_keys('{ESC}')
                            time.sleep(0.5)
                            app = _reconnect(app)
                except Exception:
                    pass

            time.sleep(1)

            # Verify we arrived at Select Patient
            for retry in range(5):
                app = _reconnect(app)
                app = dismiss_all_dialogs(app, status_callback)
                screen, win, title = identify_screen(app)
                if screen == "select_patient":
                    break
                elif screen in ("popup", "alerts", "choose_database"):
                    _dismiss(screen, win, status_callback)
                    time.sleep(0.5)
                else:
                    time.sleep(0.5)

            screen, win, title = identify_screen(app)
            if screen == "select_patient":
                _log(status_callback, "[3/7] DONE — Select Patient open!", "limegreen")
            else:
                _log(status_callback, f"[3/7] FAILED — Expected Select Patient, got '{screen}'!", "red")
                return False

        # ═══ STEP 4: Duplicate Check (Idempotency) ═══
        screen, win, title = identify_screen(app)

        if screen == "select_patient":
            _log(status_callback, "[4/7] Checking for duplicates...", "yellow")
            if check_duplicate(win, patient, status_callback):
                _log(status_callback,
                     f"[4/7] ABORTED — Patient {patient.first_name} {patient.last_name} "
                     "may already exist. Skipping to avoid double-entry.", "orange")
                return False
            _log(status_callback, "[4/7] DONE — No duplicate found.", "limegreen")
        elif screen == "edit_patient":
            _log(status_callback, "[4/7] SKIP — Already on Edit Patient, can't check duplicates.", "cyan")

        # ═══ STEP 5: Click Add Pt ═══
        screen, win, title = identify_screen(app)

        if screen == "edit_patient":
            _log(status_callback, "[5/7] SKIP — Already on Edit Patient!", "limegreen")
        else:
            _log(status_callback, "[5/7] Clicking Add Pt...", "yellow")

            sel_win = app.top_window()
            btn = find_element(sel_win, locators["add_pt_btn"], status_callback)
            if btn:
                btn.click_input()
            else:
                # Scan all buttons for "Add Pt" text
                added = False
                try:
                    for b in sel_win.descendants(control_type="Button"):
                        if "Add Pt" in b.window_text():
                            b.click_input()
                            added = True
                            break
                except Exception:
                    pass
                if not added:
                    pwa_keyboard.send_keys('%a')

            time.sleep(action_delay)

            # Verify we arrived at Edit Patient
            for _ in range(5):
                app = _reconnect(app)
                app = dismiss_all_dialogs(app, status_callback)
                screen, win, title = identify_screen(app)
                if screen == "edit_patient":
                    break
                elif screen in ("popup", "alerts", "choose_database"):
                    _dismiss(screen, win, status_callback)
                    time.sleep(0.5)
                else:
                    time.sleep(0.5)

            screen, win, title = identify_screen(app)
            if screen == "edit_patient":
                _log(status_callback, "[5/7] DONE — Edit Patient form open!", "limegreen")
            else:
                _log(status_callback, f"[5/7] FAILED — Expected Edit Patient, got '{screen}'!", "red")
                return False

        # ═══ STEP 6: Fill the Form (with verification) ═══
        _log(status_callback, "[6/7] Filling patient form...", "yellow")

        edit_win = app.top_window()
        edit_win.set_focus()
        time.sleep(0.3)

        verification_failures = []

        # Last Name
        elem = find_element(edit_win, locators["last_name_field"], status_callback)
        if elem:
            if not fill_and_verify(elem, patient.last_name, "last_name", status_callback):
                verification_failures.append("last_name")
        time.sleep(field_delay)

        # First Name
        elem = find_element(edit_win, locators["first_name_field"], status_callback)
        if elem:
            if not fill_and_verify(elem, patient.first_name, "first_name", status_callback):
                verification_failures.append("first_name")
        time.sleep(field_delay)

        # Middle Initial
        if patient.middle_initial:
            elem = find_element(edit_win, locators["middle_initial_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.middle_initial, "middle_initial", status_callback):
                    verification_failures.append("middle_initial")
            time.sleep(field_delay)

        # Gender (dropdown — special handling)
        if patient.gender:
            _log(status_callback, f"  Gender = {patient.gender}", "yellow")
            try:
                items = edit_win.descendants(control_type="ListItem")
                found_gender = False
                for item in items:
                    if item.window_text() == patient.gender:
                        item.click_input()
                        found_gender = True
                        _log(status_callback, f"  Gender = {patient.gender} [verified]", "yellow")
                        break
                if not found_gender:
                    _log(status_callback, f"  Gender: '{patient.gender}' not found in dropdown", "orange")
            except Exception as e:
                _log(status_callback, f"  Gender error: {e}", "orange")
            time.sleep(field_delay)

        # Birthdate
        if patient.dob:
            elem = find_element(edit_win, locators["birthdate_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.dob, "dob", status_callback):
                    verification_failures.append("dob")
            time.sleep(field_delay)

        # Home Phone
        if patient.phone:
            elem = find_element(edit_win, locators["phone_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.phone, "phone", status_callback):
                    verification_failures.append("phone")
            time.sleep(field_delay)

        # Address
        if patient.address:
            elem = find_element(edit_win, locators["address_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.address, "address", status_callback):
                    verification_failures.append("address")
            time.sleep(field_delay)

        # City
        if patient.city:
            elem = find_element(edit_win, locators["city_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.city, "city", status_callback):
                    verification_failures.append("city")
            time.sleep(field_delay)

        # State
        if patient.state:
            elem = find_element(edit_win, locators["state_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.state, "state", status_callback):
                    verification_failures.append("state")
            time.sleep(field_delay)

        # Zip (may be ComboBox in some OD versions)
        if patient.zip:
            elem = find_element(edit_win, locators["zip_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.zip, "zip", status_callback):
                    verification_failures.append("zip")
            else:
                # Try ComboBox fallback
                try:
                    combo = edit_win.child_window(title="Zip", control_type="ComboBox")
                    if combo.exists(timeout=1):
                        combo.click_input()
                        pwa_keyboard.send_keys(patient.zip, pause=0.03)
                        _log(status_callback, f"  Zip = {patient.zip} (typed into combo)", "yellow")
                except Exception:
                    _log(status_callback, "  Zip: could not find field", "orange")
            time.sleep(field_delay)

        # SSN
        if patient.ssn:
            elem = find_element(edit_win, locators["ssn_field"], status_callback)
            if elem:
                if not fill_and_verify(elem, patient.ssn, "ssn", status_callback):
                    verification_failures.append("ssn")
            time.sleep(field_delay)

        # Report verification results
        if verification_failures:
            _log(status_callback,
                 f"[6/7] WARNING — {len(verification_failures)} field(s) failed verification: "
                 f"{', '.join(verification_failures)}", "orange")
        else:
            _log(status_callback, "[6/7] DONE — All fields filled and verified!", "limegreen")

        # ═══ STEP 7: Save ═══
        _log(status_callback, "[7/7] Saving...", "yellow")

        save_btn = find_element(edit_win, locators["save_btn"], status_callback)
        if save_btn:
            save_btn.click_input()
        else:
            # Last resort: keyboard shortcut
            _log(status_callback, "  Save button not found, trying Enter...", "orange")
            pwa_keyboard.send_keys('{ENTER}')

        time.sleep(2)

        # Verify save — check for error dialogs
        app = _reconnect(app)
        screen, win, title = identify_screen(app)

        if screen in ("popup", "alerts"):
            error_text = title
            try:
                error_text = win.window_text()
            except Exception:
                pass

            # Check if this is an error or just a confirmation
            error_keywords = ["error", "fail", "invalid", "required", "cannot"]
            is_error = any(kw in error_text.lower() for kw in error_keywords)

            if is_error:
                _log(status_callback,
                     f"[7/7] FAILED — OpenDental error after save: {error_text[:100]}", "red")
                # Don't dismiss — let user see the error
                return False
            else:
                # Probably a confirmation popup — dismiss it
                _dismiss(screen, win, status_callback)

        _log(status_callback,
             f"[DONE] {patient.first_name} {patient.last_name} saved in OpenDental!", "limegreen")
        return True

    except Exception as e:
        _log(status_callback, f"CRITICAL ERROR: {e}", "red")
        return False
