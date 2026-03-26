"""Browser automation engine using Playwright.
Works on Mac, Windows, and Linux — fills patient forms via real browser."""

import os
import time
from core.patient import Patient, FIELD_ORDER, GENDER_MAP


def run_browser_automation(patient, status_callback, form_url=None, headless=False):
    """Open a browser, navigate to the patient form, and fill in all fields.

    Args:
        patient: Patient dataclass with field values
        status_callback: function(text, color) for status updates
        form_url: URL of the form to fill. Defaults to local test_form.html
        headless: If True, run browser without visible window
    """
    from playwright.sync_api import sync_playwright

    # Default to local test form
    if not form_url:
        test_form = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_form.html")
        form_url = f"file://{os.path.abspath(test_form)}"

    status_callback("Launching browser...", "yellow")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()

        status_callback("Navigating to patient form...", "yellow")
        page.goto(form_url)
        page.wait_for_load_state("domcontentloaded")
        time.sleep(1)

        errors = []

        for field_name in FIELD_ORDER:
            value = getattr(patient, field_name, "")
            if not value:
                continue

            display = "***" if field_name == "ssn" else value
            status_callback(f"Filling {field_name}: {display}", "yellow")

            try:
                selector = f"[data-field='{field_name}']"
                element = page.query_selector(selector)

                if not element:
                    # Fallback: try by id or name
                    element = page.query_selector(f"#{field_name}") or page.query_selector(f"[name='{field_name}']")

                if not element:
                    errors.append(f"{field_name}: field not found on page")
                    continue

                tag = element.evaluate("el => el.tagName.toLowerCase()")

                if tag == "select":
                    page.select_option(selector, value=value)
                else:
                    element.click()
                    element.fill("")  # Clear first
                    element.type(value, delay=50)

                time.sleep(0.3)

            except Exception as e:
                errors.append(f"{field_name}: {e}")

        # Click save button
        status_callback("Clicking Save...", "yellow")
        save_btn = page.query_selector("#saveBtn") or page.query_selector("button[type='submit']")
        if save_btn:
            save_btn.click()
            time.sleep(1)

        # Check for success message
        saved_msg = page.query_selector("#savedMsg")
        if saved_msg and saved_msg.is_visible():
            status_callback(f"Patient saved: {patient.first_name} {patient.last_name}", "limegreen")
        elif errors:
            status_callback(f"Done with errors: {', '.join(errors)}", "orange")
        else:
            status_callback(f"Form filled: {patient.first_name} {patient.last_name}", "limegreen")

        # Keep browser open for 5 seconds so user can see the result
        status_callback("Browser staying open for 5s...", "cyan")
        time.sleep(5)

        browser.close()

    return errors


def run_browser_batch(patients_data, status_callback, form_url=None):
    """Process multiple patients through browser automation.

    Args:
        patients_data: list of (row_num, Patient, is_valid, errors) tuples
        status_callback: function(text, color)
        form_url: URL of the form
    """
    from playwright.sync_api import sync_playwright

    if not form_url:
        test_form = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_form.html")
        form_url = f"file://{os.path.abspath(test_form)}"

    status_callback("Launching browser for batch...", "yellow")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        total = len(patients_data)

        for row_num, patient, is_valid, val_errors in patients_data:
            if not is_valid:
                status_callback(f"Row {row_num}: skipped (invalid)", "orange")
                continue

            status_callback(f"Patient {row_num}/{total}: {patient.first_name} {patient.last_name}", "yellow")

            page.goto(form_url)
            page.wait_for_load_state("domcontentloaded")
            time.sleep(0.5)

            for field_name in FIELD_ORDER:
                value = getattr(patient, field_name, "")
                if not value:
                    continue

                try:
                    selector = f"[data-field='{field_name}']"
                    element = page.query_selector(selector)
                    if not element:
                        element = page.query_selector(f"#{field_name}")
                    if not element:
                        continue

                    tag = element.evaluate("el => el.tagName.toLowerCase()")
                    if tag == "select":
                        page.select_option(selector, value=value)
                    else:
                        element.click()
                        element.fill("")
                        element.type(value, delay=30)
                    time.sleep(0.2)
                except Exception:
                    pass

            # Save
            save_btn = page.query_selector("#saveBtn") or page.query_selector("button[type='submit']")
            if save_btn:
                save_btn.click()
                time.sleep(1)

            status_callback(f"Saved: {patient.first_name} {patient.last_name}", "limegreen")
            time.sleep(1)

        status_callback(f"Batch complete! {total} patients processed.", "limegreen")
        time.sleep(3)
        browser.close()
