# Bed Utilization System - Project Context

## Project Overview
This is a Ghana Health Service hospital bed utilization management system built with Python (openpyxl + win32com) that generates Excel workbooks with VBA for data entry and reporting.

**Hospital:** Hohoe Municipal Hospital
**Year:** 2026
**Primary Technologies:** Python, Excel VBA, openpyxl, win32com

---

## System Architecture

### Build Process (Two-Phase)
1. **Phase 1** ([phase1_structure.py](phase1_structure.py)): Creates Excel structure using openpyxl
   - Sheets: Control, DailyData, Admissions, DeathsData, TransfersData, Ward sheets, Monthly Summary, Ages Summary, etc.
   - Excel Tables (ListObjects) for all data
   - Formulas and formatting

2. **Phase 2** (vba_injection package): Injects VBA via win32com
   - Package structure: utils, ui_helpers, userform_builder, navigation, core
   - VBA Modules: modConfig, modDataAccess, modReports, modNavigation, modYearEnd
   - UserForms: Daily Entry, Record Admission, Record Death, Record Ages Entry, Ward Manager, Preferences Manager
   - Navigation buttons on Control sheet

### Configuration Files
- `wards_config.json`: Ward definitions (codes, names, bed complements, emergency status)
- `hospital_preferences.json`: User preferences (show emergency totals, subtract <24hr deaths)
- `carry_forward_2026.json`: Previous year remaining patients for each ward

### Build Command
```bash
python build_workbook.py
```

---

## Critical Design Patterns

### Date Handling
**Date Columns:**
- DailyData: EntryDate (col 1), EntryTimestamp (col 12)
- Admissions: AdmissionDate (col 2), EntryTimestamp (col 11)
- DeathsData: DateOfDeath (col 2), EntryTimestamp (col 13)
- TransfersData: TransferDate (col 2), EntryTimestamp (col 8)

**Date Formats:**
- Date columns: `yyyy-mm-dd`
- Timestamp columns: `yyyy-mm-dd hh:mm`

**Date Initialization:**
- Column formats are initialized in `initialize_date_formats()` during build
- Individual cell formats are set when data is saved via VBA
- Use `FixAllDateFormats()` procedure to fix existing date formatting issues

### Remaining Calculation System
**Critical Bug Fixed:** GetLastRemainingForWard had early exit in backward scan causing incorrect values.

**Calculation Flow:**
1. `SaveDailyEntry()` → Writes data to table
2. `CalculateRemainingForRow()` → Calculates PrevRemaining and Remaining as VALUES
3. `RecalculateSubsequentRows()` → Cascades changes forward
4. Ward sheets display data via SUMIFS formulas (read-only)

**Key Functions:**
- `GetLastRemainingForWard(wardCode, entryDate)`: Finds previous remaining for a ward/date
- `CalculateRemainingForRow(targetRow)`: Calculates PrevRemaining and Remaining
- `RecalculateAllRows()`: Manual recalculation of all rows
- `VerifyCalculations()`: Diagnostic to check calculation accuracy

### Ward Management
**Dynamic Ward System:**
- Wards defined in `wards_config.json`
- VBA reads from tblWardConfig table on Control sheet
- Ward Manager form allows add/edit/delete operations
- `ExportWardConfig` button saves current config to JSON

**Ward Properties:**
- `WardCode`: Unique identifier (e.g., "MW", "FW", "MAE")
- `WardName`: Display name (e.g., "Male Medical")
- `BedComplement`: Number of beds
- `PrevYearRemaining`: Carry-forward from previous year
- `IsEmergency`: Boolean for emergency wards
- `DisplayOrder`: Sort order in reports

---

## Data Entry Forms

### Daily Bed Entry Form
- **Purpose:** Record daily admissions, discharges, deaths, transfers for each ward
- **Key Features:**
  - Month/Day/Ward selection
  - Automatic calculation of remaining patients
  - Shows previous remaining from last entry
  - Edit mode for existing entries
  - Recent entries list for quick navigation

### Record Admission Form
- **Purpose:** Record individual patient admission details
- **Fields:** AdmissionDate, Ward, PatientID, Name, Age, AgeUnit, Sex, NHIS status
- **Auto-generation:** AdmissionID (format: YYYY-00001)

### Record Death Form
- **Purpose:** Record individual death details
- **Fields:** DateOfDeath, Ward, FolderNumber, Name, Age, Sex, CauseOfDeath, DeathWithin24Hrs
- **Auto-generation:** DeathID (format: DYYYY-00001)

### Record Ages Entry Form
- **Purpose:** Bulk age group entry (faster than individual admissions)
- **Age Groups:** <1 day, 1-6 days, 7-27 days, 28 days-<1 year, 1-4 years, 5-14 years, 15-24 years, 25-44 years, 45-64 years, 65-74 years, 75-84 years, 85+ years

---

## Reports & Analysis

### Monthly Summary
- Automatically calculated from DailyData table
- KPIs: Average Daily Bed Occupancy, Average Length of Stay, Bed Turnover Interval, Bed Turnover Rate, % Occupancy, Death Rate
- Emergency Total Remaining row (configurable via preferences)

### Deaths Report
- Populated via `RefreshDeathsReport()` from DeathsData table
- Organized by month with details: Folder Number, Date, Name, Age, Sex, Ward, NHIS

### COD Summary
- Cause of Death summary populated via `RefreshCODSummary()`
- Counts deaths by cause across all months

### Ages Summary
- Automatically calculated from Admissions table using COUNTIFS
- Breakdown: Total (M/F), Non-Insured (M/F), Insured (M/F)

### Statement of Inpatient (Yearly)
- Year-to-date summary with same KPIs as Monthly Summary

---

## Diagnostic Tools

### Fix Date Formats
**Button:** Control Sheet → "Fix Date Formats"
**VBA Procedure:** `FixAllDateFormats()`

**What it does:**
1. Applies proper date format to all date columns across all tables
2. Converts text dates to proper date values
3. Reports number of cells fixed

**When to use:**
- After importing data from other sources
- When forms show date errors on edit
- When dates appear as text or are stuck on specific values

### Recalculate All Data
**Button:** Control Sheet → "Recalculate All Data"
**VBA Procedure:** `RecalculateAllRows()`

**What it does:**
- Recalculates PrevRemaining and Remaining for all rows in DailyData
- Useful after data import or manual edits

### Verify Calculations
**Button:** Control Sheet → "Verify Calculations"
**VBA Procedure:** `VerifyCalculations()`

**What it does:**
- Checks if all PrevRemaining calculations are correct
- Reports errors to Immediate Window (Ctrl+G in VBA Editor)

---

## Rebuild & Migration

### Rebuild Workbook
**Button:** Control Sheet → "Rebuild Workbook" (Orange button)

**What it does:**
1. Exports current ward config and preferences to JSON
2. Backs up current workbook
3. Runs `build_workbook.py` to create new workbook with updated code
4. Imports all data from backup

**When to use:**
- After code changes in phase1_structure.py or phase2_vba.py
- To apply new features or bug fixes

### Import from Old Workbook
**Button:** Control Sheet → "Import from Old Workbook"

**What it does:**
- Imports data from previous workbook formats
- Recalculates all values after import

---

## OCR Tool Integration (Planned)

**Architecture:** Standalone Python tool (not embedded in VBA)
**Purpose:** Extract daily ward summary data from handwritten forms
**Technology:** TrOCR (Microsoft transformer model), OpenCV preprocessing
**Output:** CSV file for review before import to Excel
**Scope:** Summary totals only (not individual patient rows)

**Workflow:**
1. Scan handwritten daily bed utilization forms
2. Run OCR tool → generates CSV with extracted data
3. Human review via Tkinter GUI
4. Import validated CSV to Excel via VBA function

---

## Common Issues & Fixes

### Issue: Date Format Errors
**Symptoms:** "Date error" when selecting entries in forms, dates stuck on January 1
**Root Cause:** Date columns not properly formatted, dates stored as text
**Fix:**
1. Click "Fix Date Formats" button on Control sheet
2. Or manually run `FixAllDateFormats()` in VBA

### Issue: Incorrect Remaining Values
**Symptoms:** Remaining patient counts don't match expected values
**Root Cause:** Calculation errors or data out of sequence
**Fix:**
1. Click "Recalculate All Data" button
2. Click "Verify Calculations" to check for errors

### Issue: Ward Changes Not Reflecting
**Symptoms:** Added/edited wards don't show in forms or reports
**Root Cause:** Changes only in JSON, not in workbook table
**Fix:**
1. Use "Manage Wards" button to edit wards (updates both table and JSON)
2. Or use "Rebuild Workbook" after manual JSON edits

---

## Development Workflow

### Making Changes

1. **Structure Changes** (sheets, tables, formulas):
   - Edit [phase1_structure.py](phase1_structure.py)
   - Run `python build_workbook.py`

2. **VBA Changes** (modules, forms, logic):
   - Edit VBA source files in `src/vba/modules/`, `src/vba/forms/`, or `src/vba/workbook/`
   - Or edit Python injection code in `src/vba_injection/` package
   - Run `python build_workbook.py`

3. **Ward Configuration**:
   - Edit `wards_config.json` directly
   - Or use Ward Manager form in workbook
   - Rebuild workbook to apply changes

4. **Preferences**:
   - Edit `hospital_preferences.json` directly
   - Or use Preferences Manager form in workbook
   - Rebuild workbook to apply changes

### Testing Workflow

1. Make code changes
2. Close Excel completely (important for VBA injection)
3. Run `python build_workbook.py`
4. Open generated .xlsm file
5. Test functionality
6. If bugs found, edit Python code and repeat

---

## File Structure
```
bedutilization/
├── build_workbook.py           # Main build script
├── src/
│   ├── config.py               # Configuration classes
│   ├── phase1_structure.py     # Excel structure generation
│   ├── vba_injection/          # VBA injection package (refactored)
│   │   ├── __init__.py         # Package exports
│   │   ├── core.py             # Main injection logic
│   │   ├── userform_builder.py # UserForm creation
│   │   ├── ui_helpers.py       # UI control helpers
│   │   ├── navigation.py       # Navigation buttons
│   │   └── utils.py            # VBA file utilities
│   └── vba/                    # VBA source files
│       ├── modules/            # VBA standard modules
│       │   ├── modConfig.bas
│       │   ├── modDataAccess.bas
│       │   ├── modReports.bas
│       │   ├── modNavigation.bas
│       │   └── modYearEnd.bas
│       ├── forms/              # VBA UserForm code
│       │   ├── frmDailyEntry.vba
│       │   ├── frmAdmission.vba
│       │   ├── frmAgesEntry.vba
│       │   ├── frmDeath.vba
│       │   ├── frmWardManager.vba
│       │   └── frmPreferencesManager.vba
│       └── workbook/           # Workbook/Sheet event code
│           ├── ThisWorkbook.cls
│           └── Sheet_DailyData.cls
├── config/
│   ├── wards_config.json       # Ward definitions
│   └── hospital_preferences.json # Hospital preferences
├── carry_forward_2026.json     # Year-end carry forward data
├── Bed_Utilization_2026.xlsm   # Generated workbook
├── ocr_tool/                   # OCR tool (standalone)
│   ├── models/form_schema.py
│   ├── extraction/trocr_engine.py
│   ├── preprocessing/enhance.py
│   └── ...
└── docs/                       # Documentation
    ├── GEMINI.md
    ├── TODO.txt
    ├── WARD_CONFIGURATION_GUIDE.md
    └── ...
```

---

## Important Notes

### Date Handling Best Practices
- Always use Date data type in VBA (not String)
- Use `CDate()` to convert when reading from cells
- Use `IsDate()` to validate before conversion
- Apply NumberFormat when writing dates to cells
- Initialize column formats during build

### Calculation System Best Practices
- Never edit PrevRemaining or Remaining manually (calculated values)
- Always use SaveDailyEntry to add/edit data (maintains calculations)
- Sort table before calculations (by WardCode, then EntryDate)
- Use RecalculateAllRows after bulk edits or imports

### Ward Management Best Practices
- Use Ward Manager form to add/edit wards (maintains consistency)
- Export config after ward changes (backup)
- Emergency wards must have IsEmergency = TRUE
- Maintain unique WardCode values

---

## Git Workflow

**Main Branch:** `main`

**Current Status:**
- Modified: Bed_Utilization_2026new.xlsm (working file)
- Modified: plan/documentations/TODO.txt
- Deleted: Bed_Utilization_2026.xlsm (old file)

**Commit Message Style:**
- `feat:` - New features
- `fix:` - Bug fixes
- `refactor:` - Code restructuring
- `chore:` - Maintenance tasks
- `update:` - Updates to documentation or config

---

## Support & Troubleshooting

### VBA Injection Fails
**Error:** "Trust access to the VBA project object model"
**Fix:**
1. Open Excel
2. File → Options → Trust Center → Trust Center Settings
3. Macro Settings → Check "Trust access to the VBA project object model"
4. Restart Excel and rebuild

### Excel File Locked During Build
**Error:** "Open method of Workbooks class failed"
**Fix:**
1. Close all Excel windows
2. Check Task Manager for excel.exe processes and end them
3. Run build again

### Missing Dependencies
**Error:** Module not found (openpyxl, win32com)
**Fix:**
```bash
pip install openpyxl pywin32
```

---

## Future Enhancements

- [ ] Complete OCR tool integration with VBA import function
- [ ] Add data validation rules to prevent invalid entries
- [ ] Implement user authentication/audit logging
- [ ] Create mobile data entry app (Progressive Web App)
- [ ] Add automated monthly report email distribution
- [ ] Implement database backend for multi-user access
