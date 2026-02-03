"""
Bed Utilization Workbook - Configuration
Ghana Health Service - Hohoe Municipal Hospital
"""
import calendar
import json
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
class WorkbookConfig:
    year: int
    hospital_name: str = "HOHOE MUNICIPAL HOSPITAL"
    carry_forward_path: Optional[str] = None

    WARDS: List[WardDef] = field(default_factory=lambda: [
        WardDef("MW",   "Male Medical",     32, False, 1),
        WardDef("FW",   "Female Medical",   28, False, 2),
        WardDef("CW",   "Paediatric",       27, False, 3),
        WardDef("BF",   "Block F",          20, False, 4),
        WardDef("BG",   "Block G",          14, False, 5),
        WardDef("BH",   "Block H",          22, False, 6),
        WardDef("NICU", "Neonatal",         15, False, 7),
        WardDef("MAE",  "Male Emergency",   10, True,  8),
        WardDef("FAE",  "Female Emergency", 10, True,  9),
    ])

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
        if self.carry_forward_path:
            self._load_carry_forward()

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
