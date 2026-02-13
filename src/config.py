"""
Bed Utilization Workbook - Configuration
Ghana Health Service - Hohoe Municipal Hospital
"""
import calendar
import json
import os
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class WardDef:
    code: str
    name: str
    bed_complement: int
    is_emergency: bool
    display_order: int
    prev_year_remaining: int = 0


@dataclass
class HospitalPreferences:
    """Hospital-specific behavioral preferences"""
    show_emergency_total_remaining: bool = True
    subtract_deaths_under_24hrs_from_admissions: bool = False


@dataclass
class WorkbookConfig:
    year: int
    hospital_name: str = "HOHOE MUNICIPAL HOSPITAL"
    carry_forward_path: Optional[str] = None
    wards_config_path: str = "config/wards_config.json"
    preferences_path: str = "config/hospital_preferences.json"

    WARDS: List[WardDef] = field(default_factory=list)
    preferences: HospitalPreferences = field(default_factory=HospitalPreferences)

    # Age groups for the Ages Summary report
    AGE_GROUPS = [
        ("0-28",  "Days",   0,  28),
        ("1-11",  "Months", 1,  11),
        ("1-4",   "Years",  1,   4),
        ("5-9",   "Years",  5,   9),
        ("10-14", "Years",  10, 14),
        ("15-17", "Years",  15, 17),
        ("18-19", "Years",  18, 19),
        ("20-34", "Years",  20, 34),
        ("35-49", "Years",  35, 49),
        ("50-59", "Years",  50, 59),
        ("60-69", "Years",  60, 69),
        ("70+",   "Years",  70, 200),
    ]

    MONTH_NAMES = [
        "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
        "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
    ]

    def __post_init__(self):
        # Load wards configuration
        self._load_wards_config()

        # Load hospital preferences
        self._load_preferences()

        # Load carry forward data if provided
        if self.carry_forward_path:
            self._load_carry_forward()

    def _load_wards_config(self):
        """Load ward configuration from JSON file or use defaults"""
        if os.path.exists(self.wards_config_path):
            try:
                with open(self.wards_config_path, 'r') as f:
                    config_data = json.load(f)

                wards_list = config_data.get("wards", [])

                if not wards_list:
                    raise ValueError("No wards found in configuration file")

                # Validate and load wards
                for ward_data in wards_list:
                    self._validate_ward_data(ward_data)
                    ward = WardDef(
                        code=ward_data["code"],
                        name=ward_data["name"],
                        bed_complement=ward_data["bed_complement"],
                        is_emergency=ward_data["is_emergency"],
                        display_order=ward_data["display_order"]
                    )
                    self.WARDS.append(ward)

                # Sort wards by display_order
                self.WARDS.sort(key=lambda w: w.display_order)

                print(f"Loaded {len(self.WARDS)} wards from {self.wards_config_path}")

            except Exception as e:
                print(f"Error loading wards configuration: {e}")
                print("Using default ward configuration")
                self._load_default_wards()
        else:
            print(f"Ward configuration file '{self.wards_config_path}' not found")
            print("Using default ward configuration")
            self._load_default_wards()

    def _validate_ward_data(self, ward_data: dict):
        """Validate ward configuration data"""
        required_fields = ["code", "name", "bed_complement", "is_emergency", "display_order"]
        for field in required_fields:
            if field not in ward_data:
                raise ValueError(f"Missing required field '{field}' in ward configuration")

        if not isinstance(ward_data["code"], str) or not ward_data["code"].strip():
            raise ValueError("Ward code must be a non-empty string")

        if not isinstance(ward_data["name"], str) or not ward_data["name"].strip():
            raise ValueError("Ward name must be a non-empty string")

        if not isinstance(ward_data["bed_complement"], int) or ward_data["bed_complement"] < 0:
            raise ValueError("Bed complement must be a non-negative integer")

        if not isinstance(ward_data["is_emergency"], bool):
            raise ValueError("is_emergency must be a boolean (true or false)")

        if not isinstance(ward_data["display_order"], int) or ward_data["display_order"] < 1:
            raise ValueError("Display order must be a positive integer")

    def _load_default_wards(self):
        """Load default ward configuration (fallback)"""
        self.WARDS = [
            WardDef("MW",   "Male Medical",     32, False, 1),
            WardDef("FW",   "Female Medical",   28, False, 2),
            WardDef("CW",   "Paediatric",       27, False, 3),
            WardDef("BF",   "Block F",          20, False, 4),
            WardDef("BG",   "Block G",          14, False, 5),
            WardDef("BH",   "Block H",          22, False, 6),
            WardDef("NICU", "Neonatal",         15, False, 7),
            WardDef("MAE",  "Male Emergency",   10, True,  8),
            WardDef("FAE",  "Female Emergency", 10, True,  9),
        ]

    def _load_preferences(self):
        """Load hospital preferences from JSON file or use defaults"""
        if os.path.exists(self.preferences_path):
            try:
                with open(self.preferences_path, 'r') as f:
                    data = json.load(f)
                prefs = data.get("preferences", {})
                self.preferences = HospitalPreferences(
                    show_emergency_total_remaining=prefs.get("show_emergency_total_remaining", True),
                    subtract_deaths_under_24hrs_from_admissions=prefs.get("subtract_deaths_under_24hrs_from_admissions", False)
                )
                print(f"Loaded hospital preferences from {self.preferences_path}")
            except Exception as e:
                print(f"Error loading preferences: {e}. Using defaults.")
                self.preferences = HospitalPreferences()
        else:
            print("No preferences file found. Using defaults.")
            self.preferences = HospitalPreferences()

    def _load_carry_forward(self):
        with open(self.carry_forward_path) as f:
            data = json.load(f)
        wards_data = data.get("wards", data)
        for ward in self.WARDS:
            ward.prev_year_remaining = wards_data.get(ward.code, 0)

    def days_in_month(self, month: int) -> int:
        return calendar.monthrange(self.year, month)[1]

    def ward_by_code(self, code: str) -> Optional[WardDef]:
        for w in self.WARDS:
            if w.code == code:
                return w
        return None
