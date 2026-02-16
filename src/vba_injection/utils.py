"""
VBA File Utilities

Helper functions for locating and reading VBA source files.
"""
import os


def get_vba_path(filename: str, subfolder: str = "modules") -> str:
    """
    Get absolute path to a VBA source file.
    
    Args:
        filename: Name of the file (e.g., "modConfig.bas")
        subfolder: Subfolder in src/vba/ (default: "modules")
        
    Returns:
        Absolute path to the VBA source file
        
    Example:
        >>> get_vba_path("modConfig.bas", "modules")
        'C:/path/to/src/vba/modules/modConfig.bas'
    """
    # This script is in src/vba_injection/utils.py, so go up to src/
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    vba_dir = os.path.join(current_dir, "vba", subfolder)
    return os.path.join(vba_dir, filename)


def read_vba_file(path: str) -> str:
    """
    Read content of a VBA source file.
    
    Args:
        path: Absolute path to the VBA file
        
    Returns:
        Content of the VBA file as string
        
    Raises:
        FileNotFoundError: If the VBA source file does not exist
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"VBA source file not found: {path}")
    
    with open(path, "r", encoding="utf-8") as f:
        return f.read()
