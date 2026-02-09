# Changelog - Bed Utilization System

## [Latest Update - Ward Management System] - 2026-02-07

### 🎉 Major New Feature: VBA Ward Management Form

#### Added
- **Ward Management Form (frmWardManager)**
  - User-friendly form to add, edit, and delete wards
  - No more manual JSON editing required!
  - Built-in validation to prevent errors
  - Real-time preview of all wards

- **Export Ward Config Function**
  - New "Export Ward Config" button on Control sheet
  - Saves current ward configuration to `wards_config.json`
  - Prepares configuration for next workbook rebuild

- **Dynamic Ward Loading in VBA**
  - All VBA forms now read wards from tblWardConfig table
  - No more hardcoded ward lists
  - Changes reflect immediately in forms

- **New Control Sheet Buttons**
  - "Manage Wards" - Opens the ward management form
  - "Export Ward Config" - Exports configuration to JSON

#### Changed
- **VBA modConfig Module**
  - GetWardCodes() now reads from tblWardConfig table
  - GetWardNames() now reads from tblWardConfig table
  - Added GetWardCount() function
  - Added GetWardByCode() function for ward lookups

## [Update - Simplified Ward Sheets] - 2026-02-07

### Fixed
- **Removed Malaria Cases column** from all ward entry sheets
  - The Malaria Cases column is no longer displayed on individual ward sheets
  - Simplified the ward data entry form (now 8 columns instead of 9)
  - Updated all formulas and totals to exclude Malaria data

### Added
- **User-Manageable Ward Configuration System**
  - New `wards_config.json` file for easy ward management
  - Users can now add, update, or remove wards without editing Python code
  - Automatic validation of ward configuration data
  - Comprehensive documentation in `WARD_CONFIGURATION_GUIDE.md`

### Changed
- **Ward Configuration Loading**
  - Ward definitions now loaded from external JSON file
  - Falls back to default wards if configuration file is missing or invalid
  - Added validation for all ward fields
  - Wards automatically sorted by display_order

### Benefits
1. **Easier Ward Management**: Hospital staff can add new wards by editing a simple JSON file
2. **No Code Changes Required**: Ward configuration changes don't require Python programming knowledge
3. **Validation & Error Handling**: System validates ward data and provides clear error messages
4. **Flexibility**: Each hospital can customize ward configuration to match their structure
5. **Cleaner Data Entry**: Removed unnecessary Malaria column reduces clutter

### Files Modified
- `phase1_structure.py` - Removed Malaria column from ward sheets
- `config.py` - Added ward configuration loading from JSON
- `wards_config.json` - New ward configuration file (user-editable)
- `WARD_CONFIGURATION_GUIDE.md` - Complete documentation for ward management

### Migration Notes
- Existing workbooks will continue to function normally
- To use the new ward configuration system, simply rebuild the workbook using:
  ```bash
  python build_workbook.py --year 2026
  ```
- To add new wards, edit `wards_config.json` and rebuild

### Backward Compatibility
- If `wards_config.json` is not found, the system uses the original default wards
- All existing VBA code remains unchanged
- Data entry processes remain the same (minus the Malaria column)
