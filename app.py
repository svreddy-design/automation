import customtkinter as ctk
import pyautogui
import time
import threading
import os
import platform
import subprocess
import json
from datetime import datetime
from tkinter import filedialog

from core.patient import Patient
from core.opendental import (
    load_timing, find_opendental, launch_app,
    navigate_to_add_patient, enter_patient_fields,
    enter_patient_fields_plain, save_patient, dry_run
)
from core.csv_import import read_patients_csv, load_batch_log, write_batch_log_entry

# Securely import pywinauto ONLY if the OS is Windows
try:
    if platform.system() == "Windows":
        from pywinauto import Application
except ImportError:
    pass

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")


class LegacyAutomationBot(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Practice Management Bot PRO")
        self.geometry("650x950")
        self.resizable(False, False)

        self.config_file = "config.json"
        self.app_path = self.load_config()
        self.timing = load_timing(self.config_file)

        # Auto-detect OpenDental on Windows if no path set
        if self.app_path == "winword" and platform.system() == "Windows":
            detected = find_opendental()
            if detected:
                self.app_path = detected
                self.save_config()

        # --- Header ---
        self.header_label = ctk.CTkLabel(
            self, text="Practice Management Bot PRO",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.header_label.pack(pady=(12, 2))

        os_name = platform.system()
        self.instructions = ctk.CTkLabel(
            self,
            text=f"Automates dental practice software  |  OS: {os_name}",
            text_color="gray"
        )
        self.instructions.pack(pady=(0, 8))

        # --- Scrollable Patient Form ---
        self.form_frame = ctk.CTkScrollableFrame(self, height=320)
        self.form_frame.pack(pady=5, padx=20, fill="x")

        self.entries = {}
        row = 0

        # Last Name
        ctk.CTkLabel(self.form_frame, text="Last Name *").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["last_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["last_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        self.entries["last_name"].insert(0, "Doe")
        row += 1

        # First Name
        ctk.CTkLabel(self.form_frame, text="First Name *").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["first_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["first_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        self.entries["first_name"].insert(0, "John")
        row += 1

        # Middle Initial
        ctk.CTkLabel(self.form_frame, text="Middle Initial").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["middle_initial"] = ctk.CTkEntry(self.form_frame, width=50)
        self.entries["middle_initial"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # Preferred Name
        ctk.CTkLabel(self.form_frame, text="Preferred Name").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["preferred_name"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["preferred_name"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        row += 1

        # Gender (dropdown)
        ctk.CTkLabel(self.form_frame, text="Gender").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.gender_var = ctk.StringVar(value="")
        self.entries["gender"] = ctk.CTkOptionMenu(
            self.form_frame, values=["", "Male", "Female", "Unknown"],
            variable=self.gender_var, width=150
        )
        self.entries["gender"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # DOB
        ctk.CTkLabel(self.form_frame, text="Date of Birth").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["dob"] = ctk.CTkEntry(self.form_frame, width=150, placeholder_text="MM/DD/YYYY")
        self.entries["dob"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # SSN (masked)
        ctk.CTkLabel(self.form_frame, text="SSN").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["ssn"] = ctk.CTkEntry(self.form_frame, width=150, show="*")
        self.entries["ssn"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # Address
        ctk.CTkLabel(self.form_frame, text="Address").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["address"] = ctk.CTkEntry(self.form_frame, width=300)
        self.entries["address"].grid(row=row, column=1, padx=8, pady=5, columnspan=3, sticky="w")
        row += 1

        # City
        ctk.CTkLabel(self.form_frame, text="City").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["city"] = ctk.CTkEntry(self.form_frame, width=200)
        self.entries["city"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        row += 1

        # State + Zip on same row
        ctk.CTkLabel(self.form_frame, text="State").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["state"] = ctk.CTkEntry(self.form_frame, width=50)
        self.entries["state"].grid(row=row, column=1, padx=8, pady=5, sticky="w")
        ctk.CTkLabel(self.form_frame, text="Zip").grid(row=row, column=2, padx=8, pady=5, sticky="w")
        self.entries["zip"] = ctk.CTkEntry(self.form_frame, width=80)
        self.entries["zip"].grid(row=row, column=3, padx=8, pady=5, sticky="w")
        row += 1

        # Phone
        ctk.CTkLabel(self.form_frame, text="Phone").grid(row=row, column=0, padx=8, pady=5, sticky="w")
        self.entries["phone"] = ctk.CTkEntry(self.form_frame, width=150, placeholder_text="10 digits")
        self.entries["phone"].grid(row=row, column=1, padx=8, pady=5, sticky="w")

        # --- CSV Import + Status ---
        self.mid_frame = ctk.CTkFrame(self)
        self.mid_frame.pack(pady=5, padx=20, fill="x")

        self.csv_btn = ctk.CTkButton(
            self.mid_frame, text="Import CSV",
            command=self.import_csv, fg_color="#444", hover_color="#555", width=120
        )
        self.csv_btn.pack(side="left", padx=10, pady=10)

        self.status_label = ctk.CTkLabel(
            self.mid_frame, text="Status: Ready",
            font=ctk.CTkFont(size=13), text_color="limegreen"
        )
        self.status_label.pack(side="left", padx=10, pady=10, expand=True)

        # --- Settings ---
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.pack(pady=5, padx=20, fill="x")

        self.path_label = ctk.CTkLabel(
            self.settings_frame, text="Target Application Path:",
            font=ctk.CTkFont(weight="bold")
        )
        self.path_label.pack(pady=(8, 0), padx=10)

        self.path_display = ctk.CTkEntry(self.settings_frame, width=400)
        self.path_display.pack(pady=4, padx=10)
        self.path_display.insert(0, self.app_path)
        self.path_display.configure(state="disabled")

        self.browse_btn = ctk.CTkButton(
            self.settings_frame, text="Select App Executable",
            command=self.browse_app_path, fg_color="#444", hover_color="#555"
        )
        self.browse_btn.pack(pady=(0, 8))

        # --- Action Buttons (3 modes) ---
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=(5, 5), padx=20, fill="x")

        # Row 1: Main automation engines
        if platform.system() == "Windows":
            self.pywin_btn = ctk.CTkButton(
                self.button_frame,
                text="Run PyWinAuto\n(Windows Native)",
                font=ctk.CTkFont(size=13, weight="bold"), height=50,
                command=self.start_pywinauto_thread,
                fg_color="#005b96", hover_color="#03396c"
            )
            self.pywin_btn.pack(side="left", expand=True, fill="x", padx=3)

        self.pyauto_btn = ctk.CTkButton(
            self.button_frame,
            text="Run PyAutoGUI\n(Cross-Platform)",
            font=ctk.CTkFont(size=13, weight="bold"), height=50,
            command=self.start_pyautogui_thread,
            fg_color="#b33939", hover_color="#cd6133"
        )
        self.pyauto_btn.pack(side="left", expand=True, fill="x", padx=3)

        # Row 2: Test/Demo modes
        self.test_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.test_frame.pack(pady=(0, 5), padx=20, fill="x")

        self.demo_btn = ctk.CTkButton(
            self.test_frame,
            text="Demo Mode\n(Type into TextEdit/Notepad)",
            font=ctk.CTkFont(size=12, weight="bold"), height=45,
            command=self.start_demo_thread,
            fg_color="#2d6a4f", hover_color="#40916c"
        )
        self.demo_btn.pack(side="left", expand=True, fill="x", padx=3)

        self.dryrun_btn = ctk.CTkButton(
            self.test_frame,
            text="Test / Dry Run\n(No app interaction)",
            font=ctk.CTkFont(size=12, weight="bold"), height=45,
            command=self.start_dryrun_thread,
            fg_color="#6c757d", hover_color="#868e96"
        )
        self.dryrun_btn.pack(side="left", expand=True, fill="x", padx=3)

    # ---------- Helpers ----------
    def update_status(self, text, color="limegreen"):
        self.after(0, lambda: self.status_label.configure(
            text=f"Status: {text}", text_color=color))

    def _all_buttons(self):
        btns = [self.pyauto_btn, self.browse_btn, self.csv_btn, self.demo_btn, self.dryrun_btn]
        if platform.system() == "Windows":
            btns.append(self.pywin_btn)
        return btns

    def disable_buttons(self):
        for btn in self._all_buttons():
            self.after(0, lambda b=btn: b.configure(state="disabled"))

    def enable_buttons(self):
        for btn in self._all_buttons():
            self.after(0, lambda b=btn: b.configure(state="normal"))

    def get_patient_from_gui(self):
        """Collect all form fields into a Patient object."""
        vals = {}
        for key, widget in self.entries.items():
            if key == "gender":
                vals[key] = self.gender_var.get()
            else:
                vals[key] = widget.get()
        return Patient(**vals)

    # ---------- Config ----------
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
                    return data.get("app_path", "winword")
            except (json.JSONDecodeError, IOError):
                return "winword"
        return "winword"

    def save_config(self):
        data = {}
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        data["app_path"] = self.app_path
        with open(self.config_file, 'w') as f:
            json.dump(data, f, indent=2)

    def browse_app_path(self):
        if platform.system() == "Darwin":
            file_path = filedialog.askopenfilename(
                title="Select Application",
                filetypes=[("Applications", "*.app"), ("All Files", "*.*")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title="Select Application Executable",
                filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
            )
        if file_path:
            self.app_path = file_path
            self.path_display.configure(state="normal")
            self.path_display.delete(0, "end")
            self.path_display.insert(0, self.app_path)
            self.path_display.configure(state="disabled")
            self.save_config()
            self.update_status(f"App Path Updated: {os.path.basename(file_path)}", "cyan")

    # ---------- CSV Import ----------
    def import_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Patient CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.disable_buttons()
            threading.Thread(
                target=self.run_csv_batch, args=(file_path,), daemon=True
            ).start()

    def run_csv_batch(self, csv_path):
        """Process CSV file in batch mode.
        OpenDental must already be running for real mode."""
        try:
            rows = read_patients_csv(csv_path)
            log_path = os.path.join(os.path.dirname(csv_path), "batch_log.csv")
            completed = load_batch_log(log_path)
            total = len(rows)

            for row_num, patient, is_valid, errors in rows:
                if row_num in completed:
                    self.update_status(f"Skipping row {row_num} (already done)", "cyan")
                    continue
                if not is_valid:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "skipped", "; ".join(errors)
                    )
                    self.update_status(f"Row {row_num}: skipped (missing required)", "orange")
                    continue

                self.update_status(
                    f"Patient {row_num}/{total}: {patient.first_name} {patient.last_name}",
                    "yellow"
                )
                try:
                    self._run_opendental_entry(patient)
                    status = "partial" if errors else "success"
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, status,
                        "; ".join(errors) if errors else ""
                    )
                except Exception as e:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "error", str(e)
                    )
                    self.update_status(f"Row {row_num} error: {e}", "red")

                time.sleep(self.timing["batch_patient_delay_s"])

            self.update_status(f"Batch complete! {total} patients processed.", "limegreen")
        except Exception as e:
            self.update_status(f"CSV Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- Shared OpenDental Entry ----------
    def _run_opendental_entry(self, patient):
        """Enter a single patient into OpenDental (assumes app is focused)."""
        navigate_to_add_patient(self.update_status, self.timing)
        field_errors = enter_patient_fields(patient, self.update_status, self.timing)
        save_patient(self.update_status)
        if field_errors:
            self.update_status(
                f"Saved with warnings: {', '.join(field_errors)}", "orange"
            )
        else:
            name = f"{patient.first_name} {patient.last_name}"
            self.update_status(f"Patient Saved: {name}", "limegreen")

    # ---------- Test / Dry Run Mode ----------
    def start_dryrun_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_dryrun, daemon=True).start()

    def run_dryrun(self):
        """Simulate automation without touching any app. Works on any OS."""
        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                return
            dry_run(patient, self.update_status)
        except Exception as e:
            self.update_status(f"Dry Run Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- Demo Mode (TextEdit/Notepad) ----------
    def start_demo_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_demo, daemon=True).start()

    def run_demo(self):
        """Open a text editor and type patient data into it. Proves automation works on any OS."""
        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                return

            self.update_status("Opening text editor for demo...", "yellow")

            if platform.system() == "Darwin":
                os.system("open -a TextEdit")
                time.sleep(2)
                os.system("osascript -e 'tell application \"TextEdit\" to activate'")
                # Create new doc
                pyautogui.hotkey('command', 'n')
                time.sleep(1)
            elif platform.system() == "Windows":
                subprocess.Popen(["notepad.exe"])
                time.sleep(2)
            else:
                # Linux
                for editor in ["gedit", "xed", "mousepad", "nano"]:
                    try:
                        subprocess.Popen([editor])
                        time.sleep(2)
                        break
                    except FileNotFoundError:
                        continue

            self.update_status("Typing patient data into editor...", "yellow")
            enter_patient_fields_plain(patient, self.update_status, self.timing)

            self.update_status("Demo complete! Check the text editor.", "limegreen")
        except Exception as e:
            self.update_status(f"Demo Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- PyWinAuto Engine (Windows Only) ----------
    def start_pywinauto_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pywinauto, daemon=True).start()

    def run_pywinauto(self):
        if platform.system() != "Windows":
            self.update_status("Error: PyWinAuto ONLY works on Windows!", "red")
            self.enable_buttons()
            return

        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                self.enable_buttons()
                return

            app_name = os.path.basename(self.app_path)

            if not launch_app(self.app_path, self.update_status, self.timing):
                self.enable_buttons()
                return

            # Connect via PyWinAuto UIA
            self.update_status("Connecting to application...", "yellow")
            try:
                app = Application(backend="uia").connect(path=self.app_path, timeout=15)
            except Exception:
                app = Application(backend="uia").connect(
                    title_re=f".*{app_name.split('.')[0]}.*", timeout=15
                )

            main_window = app.top_window()
            main_window.set_focus()

            if "opendental" in app_name.lower():
                self._run_opendental_entry(patient)
            else:
                self._run_word_automation(app, patient)

        except Exception as e:
            self.update_status(f"Execution Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- PyAutoGUI Engine (Cross-Platform) ----------
    def start_pyautogui_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_pyautogui, daemon=True).start()

    def run_pyautogui(self):
        try:
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.update_status(f"Validation failed: {'; '.join(errors)}", "red")
                self.enable_buttons()
                return

            app_name = os.path.basename(self.app_path)

            if not launch_app(self.app_path, self.update_status, self.timing):
                self.enable_buttons()
                return

            if "opendental" in app_name.lower():
                self._run_opendental_entry(patient)
            else:
                self.update_status("Non-OpenDental mode: typing patient data...", "orange")
                enter_patient_fields_plain(patient, self.update_status, self.timing)
                self.update_status("Automation Complete!", "limegreen")

        except Exception as e:
            self.update_status(f"PyAutoGUI Error: {e}", "red")
        finally:
            self.enable_buttons()

    # ---------- Word Fallback (Windows only) ----------
    def _run_word_automation(self, app, patient):
        self.update_status("Selecting 'Blank Document'...", "yellow")
        pyautogui.press('enter')
        time.sleep(4)

        doc_window = app.top_window()
        doc_window.set_focus()
        time.sleep(1)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        text = (
            f"Automated Patient Entry (Pro Mode)\n"
            f"First Name: {patient.first_name}\n"
            f"Last Name: {patient.last_name}\n"
            f"Entry Time: {timestamp}"
        )
        pyautogui.write(text, interval=0.05)
        time.sleep(1)

        pyautogui.press('f12')
        time.sleep(2)
        safe_fn = "".join(x for x in patient.first_name if x.isalnum())
        safe_ln = "".join(x for x in patient.last_name if x.isalnum())
        pyautogui.write(f"Patient_{safe_fn}_{safe_ln}_{timestamp}")
        time.sleep(1)
        pyautogui.press('enter')
        self.update_status("Document saved!", "limegreen")


if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    app = LegacyAutomationBot()
    app.mainloop()
