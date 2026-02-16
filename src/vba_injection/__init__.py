"""
VBA Injection Package for Bed Utilization System

This package handles the injection of VBA code into Excel workbooks using win32com.
It creates UserForms, standard modules, navigation buttons, and saves as .xlsm format.
"""

from .core import inject_vba, initialize_date_formats

__all__ = ["inject_vba", "initialize_date_formats"]
__version__ = "2.0.0"
