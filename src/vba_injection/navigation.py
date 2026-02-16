"""
Navigation Button Creation

Functions for creating navigation buttons on the Control sheet.
"""
from typing import Any
from .ui_helpers import add_sheet_button


def create_nav_buttons(wb: Any) -> None:
    """
    Add navigation shape-buttons to the Control sheet.
    
    Creates styled buttons for:
    - Data entry forms (Daily Entry, Admission, Death, Ages Entry)
    - Reports and management (Refresh Reports, Manage Wards, Preferences)
    - Configuration (Export Ward Config, Export Year-End)
    - Rebuild and diagnostics
    
    Args:
        wb: Excel Workbook object
    """
    ws = wb.Sheets("Control")

    # Remove placeholder text (now the cells themselves will have the text)
    for row in [9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35]:  # All button rows
        ws.Range(f"A{row}:C{row}").ClearContents

    # Set cell values which will become button captions
    ws.Range("A9").Value = "Daily Bed Entry"
    ws.Range("A11").Value = "Record Admission"
    ws.Range("A13").Value = "Record Death"
    ws.Range("A15").Value = "Record Ages Entry"
    ws.Range("A17").Value = "Refresh Reports"
    ws.Range("A19").Value = "Manage Wards"
    ws.Range("A21").Value = "Export Ward Config"
    ws.Range("A23").Value = "Export Year-End"

    # Create buttons with macros
    add_sheet_button(ws, "btnDailyEntry", "Control!A9:C9", "ShowDailyEntry")
    add_sheet_button(ws, "btnAdmission", "Control!A11:C11", "ShowAdmission")
    add_sheet_button(ws, "btnDeath", "Control!A13:C13", "ShowDeath")
    add_sheet_button(ws, "btnAgesEntry", "Control!A15:C15", "ShowAgesEntry")
    add_sheet_button(ws, "btnRefresh", "Control!A17:C17", "ShowRefreshReports")
    add_sheet_button(ws, "btnManageWards", "Control!A19:C19", "ShowWardManager")
    add_sheet_button(ws, "btnExportConfig", "Control!A21:C21", "ExportWardConfig")
    add_sheet_button(ws, "btnExportYearEnd", "Control!A23:C23", "ExportCarryForward")
    add_sheet_button(ws, "btnPreferences", "Control!A25:C25", "ShowPreferencesInfo")

    # Rebuild button (special orange button)
    add_sheet_button(ws, "btnRebuild", "Control!A27:C27", "RebuildWorkbookWithPreferences")

    # Diagnostic buttons (row 29, 31, 33 for spacing)
    ws.Range("A29").Value = "Import from Old Workbook"
    ws.Range("A31").Value = "Recalculate All Data"
    ws.Range("A33").Value = "Verify Calculations"
    add_sheet_button(ws, "btnImport", "Control!A29:C29", "ImportFromOldWorkbook")
    add_sheet_button(ws, "btnRecalcAll", "Control!A31:C31", "RecalculateAllRows")
    add_sheet_button(ws, "btnVerify", "Control!A33:C33", "VerifyCalculations")
    # Note: "Fix Date Formats" button removed - date formats now initialized automatically during build
