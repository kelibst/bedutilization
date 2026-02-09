# Ward Configuration Guide

## Overview

The Bed Utilization System now supports **user-manageable ward configurations**. You can add, update, or remove wards without modifying the Python source code.

## Configuration File

**File Location:** `wards_config.json` (in the project root directory)

## How to Add a New Ward

1. **Open** `wards_config.json` in any text editor

2. **Copy** an existing ward entry (the entire block between `{` and `}`)

3. **Paste** it at the end of the `"wards"` array (before the closing `]`)

4. **Modify** the values:
   ```json
   {
     "code": "UNIQUE_CODE",
     "name": "Ward Full Name",
     "bed_complement": 20,
     "is_emergency": false,
     "display_order": 10
   }
   ```

5. **Add a comma** after the previous ward entry if needed

6. **Save** the file

7. **Rebuild** the workbook:
   ```bash
   python build_workbook.py --year 2026
   ```

## Field Descriptions

| Field | Type | Description | Example |
|-------|------|-------------|---------|
| `code` | String | Unique identifier for the ward (2-10 characters) | `"MW"`, `"NICU"`, `"BF"` |
| `name` | String | Full ward name (shown in reports) | `"Male Medical"`, `"Neonatal"` |
| `bed_complement` | Integer | Total number of beds in the ward | `32`, `15`, `10` |
| `is_emergency` | Boolean | `true` if emergency ward, `false` otherwise | `true`, `false` |
| `display_order` | Integer | Order in reports (1, 2, 3, etc.) | `1`, `2`, `10` |

## Example: Adding a New Ward

### Before:
```json
{
  "code": "FAE",
  "name": "Female Emergency",
  "bed_complement": 10,
  "is_emergency": true,
  "display_order": 9
}
```

### After (adding ICU ward):
```json
{
  "code": "FAE",
  "name": "Female Emergency",
  "bed_complement": 10,
  "is_emergency": true,
  "display_order": 9
},
{
  "code": "ICU",
  "name": "Intensive Care Unit",
  "bed_complement": 8,
  "is_emergency": true,
  "display_order": 10
}
```

## How to Update a Ward

1. **Open** `wards_config.json`
2. **Find** the ward you want to update
3. **Modify** the values (you can change name, bed_complement, etc.)
4. **Save** the file
5. **Rebuild** the workbook

## How to Remove a Ward

1. **Open** `wards_config.json`
2. **Delete** the entire ward entry (from `{` to `}` including the comma)
3. **Save** the file
4. **Rebuild** the workbook

**⚠️ Warning:** Removing a ward will affect all reports. Only do this if you're certain the ward is no longer needed.

## Validation Rules

The system will validate your configuration when building the workbook:

- ✅ **Ward codes** must be unique and non-empty
- ✅ **Ward names** must be non-empty
- ✅ **Bed complement** must be a non-negative number (0 or greater)
- ✅ **is_emergency** must be `true` or `false` (lowercase, no quotes)
- ✅ **display_order** must be a positive number (1 or greater)

## Troubleshooting

### Error: "No wards found in configuration file"
- Check that the `"wards"` array is not empty
- Ensure the JSON syntax is correct (commas, brackets, quotes)

### Error: "Missing required field 'X' in ward configuration"
- Make sure all required fields are present: `code`, `name`, `bed_complement`, `is_emergency`, `display_order`

### Error: "Invalid JSON syntax"
- Use a JSON validator (https://jsonlint.com) to check for syntax errors
- Common issues: missing commas, extra commas, missing quotes

### System Uses Default Wards
If the configuration file cannot be loaded, the system will automatically fall back to the default wards. Check the console output when building the workbook for error messages.

## Best Practices

1. **Backup** `wards_config.json` before making changes
2. **Test** changes by rebuilding the workbook before using it in production
3. **Use descriptive ward names** that are clear to all users
4. **Keep ward codes short** but meaningful (2-6 characters recommended)
5. **Assign logical display orders** - wards appear in reports in this order

## Impact on Reports

When you add or update wards, these sheets/reports are automatically updated:

- ✅ Control sheet (ward configuration table)
- ✅ Monthly Summary (all months)
- ✅ Statement of Inpatient (yearly summary)
- ✅ Individual ward sheets (one per ward)
- ✅ All data entry forms
- ✅ All KPI calculations

## Need Help?

If you encounter issues:
1. Check the console output when running `build_workbook.py`
2. Verify your JSON syntax using an online validator
3. Review this guide for field requirements
4. Check that all required fields are present and correctly formatted
