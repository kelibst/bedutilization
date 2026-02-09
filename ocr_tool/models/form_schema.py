"""
Data models for OCR-extracted form data
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, Optional


@dataclass
class DailyWardEntry:
    """
    Extracted data structure matching tblDaily in Excel workbook
    
    This represents the summary totals from a handwritten Daily Ward State form.
    """
    # Core fields
    ward_code: str                      # "MW", "FW", "CW", etc.
    entry_date: date                    # Date from form
    admissions: int                     # Total admissions
    discharges: int                     # Total discharges
    deaths: int                         # Total deaths
    deaths_under_24: int                # Deaths within 24 hours
    transfers_in: int                   # Total transfers in
    transfers_out: int                  # Total transfers out
    
    # Validation field (from form's summary table)
    remained_midnight: Optional[int] = None  # "Remained at Midnight" from form
    
    # Metadata
    confidence_scores: Dict[str, float] = field(default_factory=dict)  # Per-field OCR confidence
    source_image: Optional[str] = None                                  # Path to source image
    reviewed_by: Optional[str] = None                                   # Who reviewed the OCR result
    notes: Optional[str] = None                                         # Manual correction notes
    
    def __post_init__(self):
        """Validate data after initialization"""
        # Ensure non-negative values
        for field_name in ['admissions', 'discharges', 'deaths', 'deaths_under_24', 
                           'transfers_in', 'transfers_out']:
            value = getattr(self, field_name)
            if value < 0:
                raise ValueError(f"{field_name} cannot be negative: {value}")
        
        # Ensure deaths_under_24 <= deaths
        if self.deaths_under_24 > self.deaths:
            raise ValueError(f"Deaths under 24hrs ({self.deaths_under_24}) cannot exceed total deaths ({self.deaths})")
    
    def calculate_expected_remaining(self, prev_remaining: int) -> int:
        """
        Calculate what the remaining should be based on the formula:
        Remaining = PrevRemaining + Admissions + TransfersIn - Discharges - Deaths - TransfersOut
        """
        return (prev_remaining + self.admissions + self.transfers_in - 
                self.discharges - self.deaths - self.transfers_out)
    
    def validate_remained_consistency(self, prev_remaining: int, tolerance: int = 2) -> tuple[bool, Optional[str]]:
        """
        Check if the remained_midnight from the form matches the calculated value
        
        Args:
            prev_remaining: Previous day's remaining count
            tolerance: Allow small discrepancies (default 2)
        
        Returns:
            (is_valid, error_message)
        """
        if self.remained_midnight is None:
            return True, None  # No validation needed if field not extracted
        
        expected = self.calculate_expected_remaining(prev_remaining)
        diff = abs(self.remained_midnight - expected)
        
        if diff > tolerance:
            return False, f"Remained mismatch: form shows {self.remained_midnight}, calculated {expected} (diff: {diff})"
        
        return True, None
    
    def get_average_confidence(self) -> float:
        """Calculate average OCR confidence across all fields"""
        if not self.confidence_scores:
            return 0.0
        return sum(self.confidence_scores.values()) / len(self.confidence_scores)
    
    def get_low_confidence_fields(self, threshold: float = 0.70) -> list[str]:
        """Get list of fields with OCR confidence below threshold"""
        return [field for field, score in self.confidence_scores.items() 
                if score < threshold]
    
    def to_csv_row(self) -> dict:
        """Convert to dictionary for CSV export"""
        return {
            'Ward': self.ward_code,
            'Date': self.entry_date.strftime('%Y-%m-%d'),
            'Admissions': self.admissions,
            'Discharges': self.discharges,
            'Deaths': self.deaths,
            'DeathsU24': self.deaths_under_24,
            'TransfersIn': self.transfers_in,
            'TransfersOut': self.transfers_out,
            'Confidence': f"{self.get_average_confidence():.2f}",
            'ReviewedBy': self.reviewed_by or '',
            'Notes': self.notes or ''
        }


@dataclass
class OCRExtractionResult:
    """
    Container for OCR extraction results including raw text and structured data
    """
    image_path: str
    success: bool
    entry: Optional[DailyWardEntry] = None
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    raw_extractions: Dict[str, tuple[str, float]] = field(default_factory=dict)  # field -> (text, confidence)
    processing_time_seconds: float = 0.0
    
    def has_errors(self) -> bool:
        """Check if extraction had any errors"""
        return len(self.errors) > 0
    
    def has_warnings(self) -> bool:
        """Check if extraction had any warnings"""
        return len(self.warnings) > 0
    
    def get_summary(self) -> str:
        """Get human-readable summary of extraction result"""
        if not self.success:
            return f"FAILED: {', '.join(self.errors)}"
        
        conf = self.entry.get_average_confidence() if self.entry else 0.0
        status = f"OK ({conf:.0%} confidence)"
        
        if self.warnings:
            status += f" - {len(self.warnings)} warning(s)"
        
        return status
