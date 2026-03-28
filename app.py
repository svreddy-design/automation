"""Practice Management Bot PRO — OpenDental Patient Entry Automation

Pure pywinauto GUI automation. Windows-only.
No database dependency — enters patients through the OpenDental UI."""

import customtkinter as ctk
import time
import threading
import os
import json
from tkinter import filedialog, messagebox

from core.patient import Patient
from core.opendental import load_timing
from core.csv_import import read_patients_csv, load_batch_log, write_batch_log_entry

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")


class PracticeManagementBot(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Practice Management Bot PRO")
        self.geometry("580x820")
        self.resizable(False, False)

        self.config_file = "config.json"
        self.config = self.load_config()
        self.timing = load_timing(self.config_file)

        # Keep bot window always on top so user can see status
        self.attributes("-topmost", True)

        # ── Header ──
        ctk.CTkLabel(
            self, text="Practice Management Bot PRO",
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(pady=(15, 0))

        ctk.CTkLabel(
            self, text="OpenDental Patient Entry Automation",
            font=ctk.CTkFont(size=13), text_color="#888"
        ).pack(pady=(2, 12))

        # ── Patient Form ──
        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.pack(pady=5, padx=20, fill="x")

        ctk.CTkLabel(
            self.form_frame, text="Patient Information",
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 5), sticky="w")

        self.entries = {}
        row = 1

        fields = [
            ("last_name", "Last Name *", 200, None, None),
            ("first_name", "First Name *", 200, None, None),
            ("middle_initial", "Middle Initial", 50, None, None),
            ("preferred_name", "Preferred Name", 200, None, None),
            ("dob", "Date of Birth", 150, "MM/DD/YYYY", None),
            ("ssn", "SSN", 150, None, "*"),
            ("address", "Address", 300, None, None),
            ("city", "City", 200, None, None),
            ("phone", "Phone", 150, "10 digits", None),
        ]

        for key, label, width, placeholder, show in fields:
            ctk.CTkLabel(self.form_frame, text=label, font=ctk.CTkFont(size=12)).grid(
                row=row, column=0, padx=(10, 5), pady=3, sticky="w"
            )
            kwargs = {"width": width}
            if placeholder:
                kwargs["placeholder_text"] = placeholder
            if show:
                kwargs["show"] = show
            entry = ctk.CTkEntry(self.form_frame, **kwargs)
            entry.grid(row=row, column=1, padx=5, pady=3, columnspan=3, sticky="w")
            self.entries[key] = entry
            row += 1

        # Gender dropdown
        ctk.CTkLabel(self.form_frame, text="Gender", font=ctk.CTkFont(size=12)).grid(
            row=row, column=0, padx=(10, 5), pady=3, sticky="w"
        )
        self.gender_var = ctk.StringVar(value="")
        self.entries["gender"] = ctk.CTkOptionMenu(
            self.form_frame, values=["", "Male", "Female", "Unknown"],
            variable=self.gender_var, width=150
        )
        self.entries["gender"].grid(row=row, column=1, padx=5, pady=3, sticky="w")
        row += 1

        # State + Zip on same row
        ctk.CTkLabel(self.form_frame, text="State", font=ctk.CTkFont(size=12)).grid(
            row=row, column=0, padx=(10, 5), pady=3, sticky="w"
        )
        self.entries["state"] = ctk.CTkEntry(self.form_frame, width=50)
        self.entries["state"].grid(row=row, column=1, padx=5, pady=3, sticky="w")
        ctk.CTkLabel(self.form_frame, text="Zip", font=ctk.CTkFont(size=12)).grid(
            row=row, column=2, padx=5, pady=3, sticky="w"
        )
        self.entries["zip"] = ctk.CTkEntry(self.form_frame, width=80)
        self.entries["zip"].grid(row=row, column=3, padx=5, pady=3, sticky="w")

        # Set defaults
        self.entries["last_name"].insert(0, "Doe")
        self.entries["first_name"].insert(0, "John")

        # ── Status Console ──
        self.status_frame = ctk.CTkFrame(self, fg_color="#0a0a1a")
        self.status_frame.pack(pady=8, padx=20, fill="x")

        ctk.CTkLabel(
            self.status_frame, text="Status Log",
            font=ctk.CTkFont(size=11, weight="bold"), text_color="#555"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.status_text = ctk.CTkTextbox(
            self.status_frame, height=100, font=ctk.CTkFont(size=12),
            fg_color="#0a0a1a", text_color="#4ecca3",
            state="disabled"
        )
        self.status_text.pack(padx=10, pady=(0, 8), fill="x")

        # ── Action Buttons ──
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=5, padx=20, fill="x")

        self.gui_btn = ctk.CTkButton(
            self.btn_frame,
            text="Enter Patient in OpenDental\n(GUI Automation)",
            font=ctk.CTkFont(size=14, weight="bold"), height=55,
            command=self.start_gui_thread,
            fg_color="#1a8a4a", hover_color="#14693a"
        )
        self.gui_btn.pack(fill="x", pady=3)

        self.csv_btn = ctk.CTkButton(
            self.btn_frame,
            text="Import CSV Batch",
            font=ctk.CTkFont(size=12), height=38,
            command=self.import_csv,
            fg_color="#444", hover_color="#555"
        )
        self.csv_btn.pack(fill="x", pady=3)

    # ── Status Logging ──
    def log(self, text, color="#4ecca3"):
        """Add a line to the status console."""
        timestamp = time.strftime("%H:%M:%S")
        self.after(0, lambda: self._append_log(f"[{timestamp}] {text}\n", color))

    def _append_log(self, text, color):
        self.status_text.configure(state="normal")
        self.status_text.insert("end", text)
        self.status_text.see("end")
        self.status_text.configure(state="disabled")

    def update_status(self, text, color="limegreen"):
        """Compatibility wrapper — also writes to log."""
        self.log(text, color)

    # ── Button State ──
    def _all_buttons(self):
        return [self.gui_btn, self.csv_btn]

    def disable_buttons(self):
        for btn in self._all_buttons():
            self.after(0, lambda b=btn: b.configure(state="disabled"))

    def enable_buttons(self):
        for btn in self._all_buttons():
            self.after(0, lambda b=btn: b.configure(state="normal"))

    # ── Get Patient ──
    def get_patient_from_gui(self):
        vals = {}
        for key, widget in self.entries.items():
            if key == "gender":
                vals[key] = self.gender_var.get()
            else:
                vals[key] = widget.get()
        return Patient(**vals)

    # ── Config ──
    def load_config(self):
        defaults = {
            "app_path": r"C:\Program Files (x86)\Open Dental\OpenDental.exe",
        }
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
                    defaults.update(data)
            except (json.JSONDecodeError, IOError):
                pass
        return defaults

    # ══════════════════════════════════════════════
    #  Human-in-the-Loop Confirmation
    # ══════════════════════════════════════════════
    def confirm_patient_entry(self, patient):
        """Show confirmation dialog before automation. Returns True if user approves."""
        msg = (
            f"Ready to enter patient into OpenDental:\n\n"
            f"  Name: {patient.first_name} {patient.last_name}\n"
            f"  Gender: {patient.gender or 'N/A'}\n"
            f"  DOB: {patient.mask_for_log('dob', patient.dob) or 'N/A'}\n"
            f"  SSN: {patient.mask_for_log('ssn', patient.ssn) or 'N/A'}\n\n"
            f"Proceed with GUI automation?"
        )
        return messagebox.askyesno("Confirm Patient Entry", msg)

    # ══════════════════════════════════════════════
    #  GUI Automation (Primary Method)
    # ══════════════════════════════════════════════
    def start_gui_thread(self):
        self.disable_buttons()
        threading.Thread(target=self.run_gui_auto, daemon=True).start()

    def run_gui_auto(self):
        try:
            self.log("── Enter Patient in OpenDental ──", "#1a8a4a")

            # Step 1: Validate
            self.log("[1] Validating patient data...")
            patient = self.get_patient_from_gui()
            is_valid, errors = patient.validate()
            if not is_valid:
                self.log(f"FAILED: {'; '.join(errors)}", "#e94560")
                return
            self.log(f"[1] Valid: {patient.first_name} {patient.last_name}")

            # Step 2: Human confirmation
            approved = self.confirm_patient_entry(patient)
            if not approved:
                self.log("Cancelled by user.", "#888")
                return

            # Step 3: Run automation
            from core.opendental_gui import automate_patient_entry

            config = dict(self.timing)
            config["app_path"] = self.config.get("app_path", "")

            success = automate_patient_entry(patient, self.update_status, config)

            if success:
                self.log("")
                self.log("Patient saved in OpenDental!", "#4ecca3")
            else:
                self.log("")
                self.log("Automation failed. Check status log for details.", "#e94560")

        except Exception as e:
            self.log(f"ERROR: {e}", "#e94560")
        finally:
            self.enable_buttons()

    # ══════════════════════════════════════════════
    #  CSV Batch Import (GUI Automation)
    # ══════════════════════════════════════════════
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
        try:
            self.log("── CSV Batch Import (GUI Automation) ──", "#f0c040")

            from core.opendental_gui import automate_patient_entry

            rows = read_patients_csv(csv_path)
            log_path = os.path.join(os.path.dirname(csv_path), "batch_log.csv")
            completed = load_batch_log(log_path)
            total = len(rows)

            config = dict(self.timing)
            config["app_path"] = self.config.get("app_path", "")
            batch_delay = self.timing.get("batch_patient_delay_s", 3)

            self.log(f"Loaded {total} patients from CSV")

            # Human confirmation for batch
            proceed = messagebox.askyesno(
                "Confirm Batch Import",
                f"Ready to enter {total} patients into OpenDental via GUI automation.\n\n"
                f"Already completed: {len(completed)}\n"
                f"Remaining: {total - len(completed)}\n\n"
                f"This will take approximately "
                f"{(total - len(completed)) * 30} seconds.\n\n"
                f"Proceed?"
            )
            if not proceed:
                self.log("Batch cancelled by user.", "#888")
                return

            success_count = 0
            skip_count = 0

            for idx, (row_num, patient, is_valid, errors) in enumerate(rows):
                if row_num in completed:
                    self.log(f"  Row {row_num}: skipped (already done)")
                    skip_count += 1
                    continue

                if not is_valid:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "skipped", "; ".join(errors)
                    )
                    self.log(f"  Row {row_num}: skipped ({'; '.join(errors)})")
                    skip_count += 1
                    continue

                self.log(f"  Patient {row_num}/{total}: {patient.first_name} {patient.last_name}...")

                success = automate_patient_entry(patient, self.update_status, config)
                if success:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "success", ""
                    )
                    success_count += 1
                else:
                    write_batch_log_entry(
                        log_path, row_num, patient.last_name,
                        patient.first_name, "error", "GUI automation failed"
                    )

                # Delay between patients (skip after last)
                if idx < len(rows) - 1:
                    time.sleep(batch_delay)

            self.log("")
            self.log(
                f"Batch complete: {success_count} saved, {skip_count} skipped, {total} total",
                "#4ecca3"
            )

        except Exception as e:
            self.log(f"CSV Error: {e}", "#e94560")
        finally:
            self.enable_buttons()


if __name__ == "__main__":
    app = PracticeManagementBot()
    app.mainloop()
