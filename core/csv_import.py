import csv
import os
from core.patient import Patient


def read_patients_csv(filepath):
    """Read patients from CSV file. Returns list of (row_number, Patient, is_valid, errors) tuples."""
    results = []
    with open(filepath, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=1):
            patient = Patient(
                last_name=row.get('last_name', '').strip(),
                first_name=row.get('first_name', '').strip(),
                middle_initial=row.get('middle_initial', '').strip(),
                preferred_name=row.get('preferred_name', '').strip(),
                gender=row.get('gender', '').strip(),
                dob=row.get('dob', '').strip(),
                ssn=row.get('ssn', '').strip(),
                address=row.get('address', '').strip(),
                city=row.get('city', '').strip(),
                state=row.get('state', '').strip(),
                zip=row.get('zip', '').strip(),
                phone=row.get('phone', '').strip(),
            )
            is_valid, errors = patient.validate()
            results.append((i, patient, is_valid, errors))
    return results


def load_batch_log(log_path):
    """Load previously completed rows from batch_log.csv. Returns set of completed row numbers."""
    completed = set()
    if not os.path.exists(log_path):
        return completed
    with open(log_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get('status') == 'success':
                completed.add(int(row['row_number']))
    return completed


def write_batch_log_entry(log_path, row_number, last_name, first_name, status, error=""):
    """Append a single entry to batch_log.csv. Never logs SSN."""
    file_exists = os.path.exists(log_path)
    with open(log_path, 'a', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['row_number', 'last_name', 'first_name', 'status', 'error'])
        writer.writerow([row_number, last_name, first_name, status, error])
