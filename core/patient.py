from dataclasses import dataclass, asdict
import re

FIELD_ORDER = [
    "last_name", "first_name", "middle_initial", "preferred_name",
    "gender", "dob", "ssn", "address", "city", "state", "zip", "phone"
]

GENDER_MAP = {"male": 0, "female": 1, "unknown": 2}


@dataclass
class Patient:
    last_name: str = ""
    first_name: str = ""
    middle_initial: str = ""
    preferred_name: str = ""
    gender: str = ""
    dob: str = ""
    ssn: str = ""
    address: str = ""
    city: str = ""
    state: str = ""
    zip: str = ""
    phone: str = ""

    def validate(self):
        """Validate fields. Returns (is_valid, errors) tuple.
        is_valid is False only if required fields are missing.
        errors list contains warnings for invalid optional fields."""
        errors = []
        if not self.last_name.strip():
            errors.append("last_name is required")
        if not self.first_name.strip():
            errors.append("first_name is required")
        if self.middle_initial and len(self.middle_initial) > 1:
            errors.append("middle_initial must be single character")
            self.middle_initial = self.middle_initial[0]
        if self.gender and self.gender.lower() not in GENDER_MAP:
            errors.append(f"gender '{self.gender}' invalid, must be Male/Female/Unknown")
            self.gender = ""
        if self.dob and not re.match(r'^\d{2}/\d{2}/\d{4}$', self.dob):
            errors.append(f"dob '{self.dob}' invalid, must be MM/DD/YYYY")
            self.dob = ""
        if self.ssn and not re.match(r'^\d{9}$', self.ssn):
            errors.append(f"ssn invalid, must be 9 digits")
            self.ssn = ""
        if self.state and not re.match(r'^[A-Z]{2}$', self.state.upper()):
            errors.append(f"state '{self.state}' invalid, must be 2-letter code")
            self.state = ""
        if self.zip and not re.match(r'^\d{5}$', self.zip):
            errors.append(f"zip '{self.zip}' invalid, must be 5 digits")
            self.zip = ""
        if self.phone and not re.match(r'^\d{10}$', self.phone):
            errors.append(f"phone '{self.phone}' invalid, must be 10 digits")
            self.phone = ""

        has_required = bool(self.last_name.strip() and self.first_name.strip())
        return has_required, errors

    def to_dict(self):
        return asdict(self)
