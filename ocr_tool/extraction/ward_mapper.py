"""
Maps ward names from handwritten forms to ward codes
"""
import sys
from pathlib import Path

# Add parent directory to path to import config
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

try:
    from config import WARDS
    WARD_CODES = {ward.code: ward.name for ward in WARDS}
    WARD_NAMES = {ward.name.upper(): ward.code for ward in WARDS}
except ImportError:
    # Fallback if config.py not available
    WARD_CODES = {
        'MW': 'Male Medical',
        'FW': 'Female Medical',
        'CW': 'Paediatric',
        'BF': 'Block F',
        'BG': 'Block G',
        'BH': 'Block H',
        'NICU': 'Neonatal',
        'MAE': 'Male Emergency',
        'FAE': 'Female Emergency'
    }
    WARD_NAMES = {name.upper(): code for code, name in WARD_CODES.items()}


def map_ward_name_to_code(ward_text: str) -> str:
    """
    Map extracted ward name text to standard ward code
    
    Args:
        ward_text: Extracted text from form (e.g., "Male", "MALE MEDICAL", "MW")
    
    Returns:
        Ward code (e.g., "MW")
    
    Raises:
        ValueError: If ward name cannot be mapped
    """
    if not ward_text:
        raise ValueError("Empty ward name")
    
    # Normalize: uppercase and strip whitespace
    normalized = ward_text.upper().strip()
    
    # Direct code match (e.g., "MW", "NICU")
    if normalized in WARD_CODES:
        return normalized
    
    # Full name match (e.g., "MALE MEDICAL")
    if normalized in WARD_NAMES:
        return WARD_NAMES[normalized]
    
    # Partial matches
    partial_matches = {
        'MALE MED': 'MW',
        'FEMALE MED': 'FW',
        'PAED': 'CW',
        'PED': 'CW',
        'CHILD': 'CW',
        'NEO': 'NICU',
        'NICU': 'NICU',
        'MALE EME': 'MAE',
        'FEM EME': 'FAE',
        'MALE': 'MW',  # Fallback for just "MALE"
        'FEMALE': 'FW'  # Fallback for just "FEMALE"
    }
    
    for keyword, code in partial_matches.items():
        if keyword in normalized:
            return code
    
    # No match found
    raise ValueError(f"Cannot map ward name: '{ward_text}' (normalized: '{normalized}')")


def get_ward_name(ward_code: str) -> str:
    """
    Get full ward name from code
    
    Args:
        ward_code: Ward code (e.g., "MW")
    
    Returns:
        Full ward name (e.g., "Male Medical")
    """
    return WARD_CODES.get(ward_code, ward_code)


def get_all_ward_codes() -> list[str]:
    """Get list of all valid ward codes"""
    return list(WARD_CODES.keys())


def is_valid_ward_code(ward_code: str) -> bool:
    """Check if ward code is valid"""
    return ward_code in WARD_CODES
