"""
Core VBA Injection Logic

Main orchestration for injecting VBA code into Excel workbooks via win32com.
Handles Excel COM interaction, module injection, and workbook saving.
"""
import os
import time
import json
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..config import WorkbookConfig

from .utils import get_vba_path, read_vba_file
from .userform_builder import (
    create_daily_entry_form,
    create_admission_form,
    create_ages_entry_form,
    create_death_form,
    create_ward_manager_form,
    create_preferences_manager_form,
)
from .navigation import create_nav_buttons


def initialize_date_formats(wb) -> None:
    """
    Initialize date column formats for all data tables.
    
    This ensures date columns are properly formatted from the start,
    preventing date format errors when editing entries in forms.
    
    Args:
        wb: Excel Workbook object
    """
    # DailyData - EntryDate (col A) and EntryTimestamp (col L)
    try:
        daily_tbl = wb.Sheets("DailyData").ListObjects("tblDaily")
        daily_tbl.ListColumns(1).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        daily_tbl.ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format DailyData columns: {e}")

    # Admissions - AdmissionDate (col B) and EntryTimestamp (col K)
    try:
        adm_tbl = wb.Sheets("Admissions").ListObjects("tblAdmissions")
        adm_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        adm_tbl.ListColumns(11).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format Admissions columns: {e}")

    # DeathsData - DateOfDeath (col B) and EntryTimestamp (col M)
    try:
        deaths_tbl = wb.Sheets("DeathsData").ListObjects("tblDeaths")
        deaths_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        deaths_tbl.ListColumns(13).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format DeathsData columns: {e}")

    # TransfersData - TransferDate (col B) and EntryTimestamp (col H)
    try:
        trans_tbl = wb.Sheets("TransfersData").ListObjects("tblTransfers")
        trans_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        trans_tbl.ListColumns(8).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format TransfersData columns: {e}")


def inject_vba(xlsx_path: str, xlsm_path: str, config: "WorkbookConfig") -> None:
    """
    Open xlsx in Excel via COM, inject VBA, save as xlsm.
    
    This is the main entry point for VBA injection. It:
    1. Opens the .xlsx file in Excel
    2. Injects VBA modules from src/vba/modules/
    3. Injects worksheet event code
    4. Creates UserForms programmatically
    5. Creates navigation buttons on Control sheet
    6. Initializes date formats
    7. Saves as .xlsm format
    
    Args:
        xlsx_path: Path to source .xlsx file
        xlsm_path: Path to output .xlsm file
        config: WorkbookConfig object with configuration settings
        
    Raises:
        FileNotFoundError: If xlsx file doesn't exist
        RuntimeError: If Excel COM interaction fails
    """
    import win32com.client

    abs_xlsx = os.path.abspath(xlsx_path)
    abs_xlsm = os.path.abspath(xlsm_path)

    # Verify xlsx file exists
    if not os.path.exists(abs_xlsx):
        raise FileNotFoundError(f"Cannot find xlsx file: {abs_xlsx}")

    print(f"Opening file: {abs_xlsx}")

    # Remove existing xlsm if it exists
    if os.path.exists(abs_xlsm):
        print(f"Removing existing xlsm: {abs_xlsm}")
        os.remove(abs_xlsm)

    # Pre-flight checks
    print("Performing pre-flight checks...")
    
    # Check if file is already open in Excel
    try:
        import psutil
        excel_processes = [p for p in psutil.process_iter(['name', 'open_files']) 
                          if p.info['name'] and 'excel' in p.info['name'].lower()]
        for proc in excel_processes:
            try:
                if proc.info['open_files']:
                    for file in proc.info['open_files']:
                        if abs_xlsx in file.path:
                            print(f"\n⚠️  WARNING: File is already open in Excel (PID: {proc.pid})")
                            print("Please close the file in Excel and try again.")
                            raise RuntimeError(f"File is already open: {abs_xlsx}")
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                pass
    except ImportError:
        # psutil not available, skip this check
        print("  (psutil not available, skipping open file check)")
    
    print("Starting Excel for VBA injection...")
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Give Excel a moment to fully initialize
        time.sleep(1)
        
        # Use standard Windows paths (backslashes) which are reliable for local files
        xlsx_path_normalized = abs_xlsx
        print(f"Opening workbook: {xlsx_path_normalized}")
        
        # Try opening with retry logic (sometimes COM needs a moment)
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                # Open with explicit parameters to avoid issues
                wb = excel.Workbooks.Open(
                    xlsx_path_normalized,
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False
                )
                print("  [OK] Workbook opened successfully")
                break
            except Exception as open_error:
                if attempt < max_retries - 1:
                    print(f"  Attempt {attempt + 1} failed, retrying in {retry_delay}s...")
                    time.sleep(retry_delay)
                else:
                    raise open_error
        
        if wb is None:
            raise RuntimeError("Failed to open workbook after all retries")
        vbproj = wb.VBProject

        # 1. Inject standard modules
        print("  Injecting VBA modules...")
        modules = [
            ("modConfig", "modConfig.bas"),
            ("modDataAccess", "modDataAccess.bas"),
            ("modReports", "modReports.bas"),
            ("modNavigation", "modNavigation.bas"),
            ("modYearEnd", "modYearEnd.bas"),
        ]
        for mod_name, filename in modules:
            module = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
            module.Name = mod_name
            code_path = get_vba_path(filename, "modules")
            module.CodeModule.AddFromString(read_vba_file(code_path))

        # 2. Inject ThisWorkbook code
        print("  Injecting ThisWorkbook code...")
        tb = vbproj.VBComponents("ThisWorkbook")
        tb_code_path = get_vba_path("ThisWorkbook.cls", "workbook")
        tb.CodeModule.AddFromString(read_vba_file(tb_code_path))

        # 2.5. Inject DailyData worksheet change event
        print("  Injecting DailyData worksheet event...")
        daily_data_injected = False
        
        # Read the event code once
        dd_code_path = get_vba_path("Sheet_DailyData.cls", "workbook")
        dd_code = read_vba_file(dd_code_path)

        for comp in vbproj.VBComponents:
            try:
                # Only check worksheet components (Type 100 = vbext_ct_Document)
                if comp.Type == 100:
                    comp_name = comp.Properties("Name").Value
                    if comp_name == "DailyData":
                        comp.CodeModule.AddFromString(dd_code)
                        print(f"    [OK] Event injected into {comp_name} worksheet")
                        daily_data_injected = True
                        break
            except Exception as e:
                # Log specific errors instead of silent failure
                print(f"    ! Component check failed: {e}")
                continue

        if not daily_data_injected:
            raise ValueError("CRITICAL: Failed to inject Worksheet_Change event into DailyData sheet!")

        # 3. Create UserForms
        print("  Creating UserForms...")
        create_daily_entry_form(vbproj)
        create_admission_form(vbproj)
        create_ages_entry_form(vbproj)
        create_death_form(vbproj)
        create_ward_manager_form(vbproj)
        create_preferences_manager_form(vbproj)

        # 4. Add navigation buttons to Control sheet
        print("  Adding navigation buttons...")
        create_nav_buttons(wb)

        # 5. Hide data sheets
        print("  Hiding data sheets...")
        wb.Sheets("DailyData").Visible = 0       # xlSheetHidden
        wb.Sheets("Admissions").Visible = 0
        wb.Sheets("DeathsData").Visible = 0
        wb.Sheets("TransfersData").Visible = 0

        # Hide individual emergency sheets by default
        try:
            wb.Sheets("Male Emergency").Visible = 0
            wb.Sheets("Female Emergency").Visible = 0
        except:
            pass


        # 5.5. Initialize date column formats
        print("  Initializing date column formats...")
        initialize_date_formats(wb)

        # 6. Save as .xlsm (FileFormat 52)
        print(f"  Saving as {abs_xlsm}...")
        wb.SaveAs(abs_xlsm, FileFormat=52)
        wb.Close(SaveChanges=False)

        print(f"Phase 2 complete: {abs_xlsm}")

    except Exception as e:
        print(f"\nERROR during VBA injection: {e}")
        print(f"Error type: {type(e).__name__}")

        # Provide specific troubleshooting based on error
        if "VBProject" in str(e) or "Programmatic access" in str(e):
            print("\nFIX: Enable VBA project access in Excel:")
            print("  1. Open Excel")
            print("  2. File > Options > Trust Center > Trust Center Settings")
            print("  3. Macro Settings > Check 'Trust access to the VBA project object model'")
            print("  4. Click OK and restart Excel")
        elif "Open" in str(e):
            print("\nPossible causes:")
            print(f"  - File exists: {os.path.exists(abs_xlsx)}")
            print(f"  - File path: {abs_xlsx}")
            print(f"  - File readable: {os.access(abs_xlsx, os.R_OK)}")
            print("\nTroubleshooting:")
            print("  1. Try opening the .xlsx file manually in Excel first")
            print("  2. Check if file is corrupted")
            print("  3. Close all Excel windows and try again")
            print("  4. Check antivirus isn't blocking file access")

        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        raise
    finally:
        if excel:
            excel.Quit()
            time.sleep(1)
