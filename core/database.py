"""Direct database insertion for OpenDental.
Connects to the MariaDB/MySQL database and inserts patients directly.
This is the most reliable method — no GUI automation needed.
Works on any platform that can reach the database."""

import datetime


def get_connection(host="localhost", port=3306, user="root", password="", database="opendental"):
    """Connect to OpenDental's MariaDB/MySQL database."""
    try:
        import mysql.connector
        conn = mysql.connector.connect(
            host=host, port=port, user=user,
            password=password, database=database
        )
        return conn
    except ImportError:
        # Try pymysql as fallback
        import pymysql
        conn = pymysql.connect(
            host=host, port=port, user=user,
            password=password, database=database
        )
        return conn


def insert_patient(patient, status_callback, db_config=None):
    """Insert a patient directly into OpenDental's database.

    Args:
        patient: Patient dataclass
        status_callback: function(text, color) for status updates
        db_config: dict with host, port, user, password, database keys
    Returns:
        patient_num (int) or None on failure
    """
    if db_config is None:
        db_config = {
            "host": "localhost",
            "port": 3306,
            "user": "root",
            "password": "",
            "database": "opendental"
        }

    status_callback("Connecting to OpenDental database...", "yellow")

    try:
        conn = get_connection(**db_config)
        cursor = conn.cursor()

        # Build the INSERT for the patient table
        # OpenDental's patient table has many columns, but only a few are needed
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d %H:%M:%S")

        # Map gender to OpenDental's PatStatus values
        # OpenDental uses: 0=Male, 1=Female, 2=Unknown
        gender_map = {"male": 0, "female": 1, "unknown": 2, "": 2}
        gender_val = gender_map.get(patient.gender.lower(), 2) if patient.gender else 2

        # Format DOB from MM/DD/YYYY to YYYY-MM-DD
        birthdate = "0001-01-01"
        if patient.dob:
            try:
                parts = patient.dob.split("/")
                if len(parts) == 3:
                    birthdate = f"{parts[2]}-{parts[0]}-{parts[1]}"
            except (ValueError, IndexError):
                pass

        # Clean phone number (remove non-digits)
        phone = "".join(c for c in patient.phone if c.isdigit()) if patient.phone else ""

        # Format phone as (XXX)XXX-XXXX if 10 digits
        phone_formatted = phone
        if len(phone) == 10:
            phone_formatted = f"({phone[:3]}){phone[3:6]}-{phone[6:]}"

        status_callback(f"Inserting {patient.first_name} {patient.last_name}...", "yellow")

        sql = """INSERT INTO patient (
            LName, FName, MiddleI, Preferred,
            Gender, Birthdate, SSN,
            Address, City, State, Zip,
            HmPhone,
            PatStatus, DateFirstVisit,
            PriProv, BillingType, SecProv
        ) VALUES (
            %s, %s, %s, %s,
            %s, %s, %s,
            %s, %s, %s, %s,
            %s,
            0, %s,
            1, 1, 0
        )"""

        values = (
            patient.last_name,
            patient.first_name,
            patient.middle_initial or "",
            patient.preferred_name or "",
            gender_val,
            birthdate,
            patient.ssn or "",
            patient.address or "",
            patient.city or "",
            patient.state or "",
            patient.zip or "",
            phone_formatted,
            date_str,
        )

        cursor.execute(sql, values)
        conn.commit()
        patient_num = cursor.lastrowid

        # Fix: OpenDental requires Guarantor = PatNum for head-of-household
        status_callback("Setting guarantor...", "yellow")
        cursor.execute(
            "UPDATE patient SET Guarantor = %s WHERE PatNum = %s",
            (patient_num, patient_num)
        )
        conn.commit()

        status_callback(f"Patient #{patient_num} saved: {patient.first_name} {patient.last_name}", "limegreen")

        cursor.close()
        conn.close()
        return patient_num

    except Exception as e:
        status_callback(f"Database Error: {e}", "red")
        return None


def insert_patients_batch(patients_data, status_callback, db_config=None):
    """Insert multiple patients from CSV data.

    Args:
        patients_data: list of (row_num, Patient, is_valid, errors) tuples
        status_callback: function(text, color)
        db_config: database connection config
    """
    total = len(patients_data)
    success = 0
    failed = 0

    for row_num, patient, is_valid, errors in patients_data:
        if not is_valid:
            status_callback(f"Row {row_num}: skipped (invalid)", "orange")
            failed += 1
            continue

        status_callback(f"Patient {row_num}/{total}: {patient.first_name} {patient.last_name}", "yellow")
        result = insert_patient(patient, status_callback, db_config)

        if result:
            success += 1
        else:
            failed += 1

    status_callback(f"Batch done! {success} saved, {failed} failed out of {total}", "limegreen")


def test_connection(status_callback, db_config=None):
    """Test if we can connect to the OpenDental database."""
    if db_config is None:
        db_config = {
            "host": "localhost",
            "port": 3306,
            "user": "root",
            "password": "",
            "database": "opendental"
        }

    status_callback("Testing database connection...", "yellow")
    try:
        conn = get_connection(**db_config)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM patient")
        count = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        status_callback(f"Connected! {count} patients in database.", "limegreen")
        return True
    except Exception as e:
        status_callback(f"Connection failed: {e}", "red")
        return False
