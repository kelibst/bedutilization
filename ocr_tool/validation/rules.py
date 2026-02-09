"""
Business rules validation for extracted form data
"""
from datetime import date, datetime
from typing import Tuple, List
import sys
from pathlib import Path

# Add parent directory to import ward mapper
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.form_schema import DailyWardEntry
from extraction.ward_mapper import is_valid_ward_code


def validate_daily_entry(
    entry: DailyWardEntry,
    current_year: int = 2026,
    prev_remaining: int = 0
) -> Tuple[List[str], List[str]]:
    """
    Validate a daily ward entry against business rules
    
    Args:
        entry: DailyWardEntry to validate
        current_year: Expected year for dates
        prev_remaining: Previous remaining count (for validation)
    
    Returns:
        (errors, warnings) - Lists of validation messages
    """
    errors = []
    warnings = []
    
    # Rule 1: Ward code must be valid
    if not is_valid_ward_code(entry.ward_code):
        errors.append(f"Invalid ward code: '{entry.ward_code}'")
    
    # Rule 2: Date validation
    if entry.entry_date is None:
        errors.append("Entry date is required")
    else:
        # Check if date is in expected year
        if entry.entry_date.year != current_year:
            warnings.append(
                f"Date year {entry.entry_date.year} differs from workbook year {current_year}"
            )
        
        # Check if date is in the future
        if entry.entry_date > date.today():
            warnings.append(
                f"Entry date {entry.entry_date} is in the future"
            )
        
        # Check if date is too old (>1 year ago)
        days_old = (date.today() - entry.entry_date).days
        if days_old > 365:
            warnings.append(
                f"Entry date is {days_old} days old - are you sure?"
            )
    
    # Rule 3: Non-negative values
    for field_name in ['admissions', 'discharges', 'deaths', 'deaths_under_24',
                       'transfers_in', 'transfers_out']:
        value = getattr(entry, field_name)
        if value < 0:
            errors.append(f"{field_name} cannot be negative: {value}")
    
    # Rule 4: Deaths under 24 hours <= Total deaths
    if entry.deaths_under_24 > entry.deaths:
        errors.append(
            f"Deaths under 24hrs ({entry.deaths_under_24}) cannot exceed "
            f"total deaths ({entry.deaths})"
        )
    
    # Rule 5: Reasonable bounds check
    if entry.admissions > 100:
        warnings.append(
            f"Very high admissions count ({entry.admissions}) - please verify"
        )
    
    if entry.discharges > 100:
        warnings.append(
            f"Very high discharges count ({entry.discharges}) - please verify"
        )
    
    if entry.deaths > 20:
        warnings.append(
            f"Very high deaths count ({entry.deaths}) - please verify"
        )
    
    # Rule 6: Remained consistency (if available)
    if entry.remained_midnight is not None and prev_remaining > 0:
        is_valid, error_msg = entry.validate_remained_consistency(prev_remaining, tolerance=2)
        if not is_valid:
            warnings.append(error_msg)
    
    # Rule 7: Total activity check (detect empty/suspicious entries)
    total_activity = (entry.admissions + entry.discharges + entry.deaths +
                     entry.transfers_in + entry.transfers_out)
    if total_activity == 0:
        warnings.append("No activity recorded - all counts are zero")
    
    return errors, warnings


def validate_date_string(date_str: str) -> Tuple[bool, str, date]:
    """
    Validate and parse date string
    
    Args:
        date_str: Date string (e.g., "12/01/2026", "2026-01-12")
    
    Returns:
        (is_valid, error_message, parsed_date)
    """
    if not date_str or not date_str.strip():
        return False, "Date is empty", None
    
    # Try common date formats
    formats = [
        "%d/%m/%Y",     # 12/01/2026
        "%m/%d/%Y",     # 01/12/2026
        "%Y-%m-%d",     # 2026-01-12
        "%d-%m-%Y",     # 12-01-2026
        "%d/%m/%y",     # 12/01/26
        "%m/%d/%y",     # 01/12/26
    ]
    
    parsed_date = None
    for fmt in formats:
        try:
            parsed_date = datetime.strptime(date_str.strip(), fmt).date()
            break
        except ValueError:
            continue
    
    if parsed_date is None:
        return False, f"Cannot parse date: '{date_str}'", None
    
    return True, "", parsed_date


def validate_integer_string(value_str: str, field_name: str = "value") -> Tuple[bool, str, int]:
    """
    Validate and parse integer string
    
    Args:
        value_str: String to parse
        field_name: Name of field (for error messages)
    
    Returns:
        (is_valid, error_message, parsed_int)
    """
    if not value_str or not value_str.strip():
        return True, "", 0  # Empty is valid (treated as 0)
    
    try:
        value = int(value_str.strip())
        if value < 0:
            return False, f"{field_name} cannot be negative: {value}", value
        return True, "", value
    except ValueError:
        return False, f"{field_name} is not a valid number: '{value_str}'", 0


def check_confidence_threshold(
    confidence_scores: dict,
    low_threshold: float = 0.70,
    medium_threshold: float = 0.85
) -> Tuple[str, List[str]]:
    """
    Check OCR confidence and categorize quality
    
    Args:
        confidence_scores: Dict of field -> confidence
        low_threshold: Below this is "low confidence"
        medium_threshold: Below this is "medium confidence"
    
    Returns:
        (overall_quality, low_confidence_fields)
        overall_quality: "high", "medium", "low"
    """
    if not confidence_scores:
        return "unknown", []
    
    avg_confidence = sum(confidence_scores.values()) / len(confidence_scores)
    low_fields = [field for field, score in confidence_scores.items()
                  if score < low_threshold]
    
    if avg_confidence >= medium_threshold and not low_fields:
        return "high", []
    elif avg_confidence >= low_threshold:
        return "medium", low_fields
    else:
        return "low", low_fields
