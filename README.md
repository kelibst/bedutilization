# Bed Utilization Management System

**Ghana Health Service - Hospital Bed Tracking & Reporting**

An automated Excel-based system for tracking daily bed utilization, patient admissions, deaths, and generating comprehensive reports for hospital management.

## 🎯 Features

- **Daily Bed Entry**: Track admissions, discharges, deaths, and transfers per ward
- **Individual Patient Records**: Record detailed admission and death information
- **Automated Reports**: Monthly summaries, age group analysis, cause of death reports
- **VBA Automation**: User-friendly forms for data entry
- **Flexible Configuration**: Easily customize ward configurations
- **Year-End Rollover**: Carry forward remaining patients to next year

## 📋 Recent Updates (2026-02-07)

### ✅ Removed Malaria Column
The Malaria Cases column has been removed from all ward entry sheets, simplifying the data entry process.

### ✅ User-Manageable Ward Configuration
You can now **add, update, or remove wards** without editing Python code!
- Edit `wards_config.json` to manage wards
- Changes automatically apply to all reports
- See [WARD_CONFIGURATION_GUIDE.md](WARD_CONFIGURATION_GUIDE.md) for details

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or later
- Microsoft Excel 2016 or later
- Python packages: `openpyxl`, `pywin32`

### Installation

1. **Install Python dependencies:**
   ```bash
   pip install openpyxl pywin32
   ```

2. **Build the workbook:**
   ```bash
   python build_workbook.py --year 2026
   ```

3. **Open the generated file:**
   - File: `Bed_Utilization_2026.xlsm`
   - Enable macros when prompted

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| [WARD_CONFIGURATION_GUIDE.md](WARD_CONFIGURATION_GUIDE.md) | Complete guide to managing ward configurations |
| [QUICK_START_WARDS.md](QUICK_START_WARDS.md) | Quick reference for common ward management tasks |
| [CHANGELOG.md](CHANGELOG.md) | History of changes and updates |

## 🏥 System Components

### Sheets Generated

1. **Control** - Main dashboard with data entry buttons
2. **DailyData** - Hidden table storing daily ward statistics
3. **Admissions** - Patient admission records
4. **DeathsData** - Individual death records
5. **TransfersData** - Patient transfer records
6. **Ward Sheets** (one per ward) - Daily bed utilization forms
7. **Monthly Summary** - Monthly statistics and KPIs for all wards
8. **Ages Summary** - Age group breakdown by month and insurance status
9. **Deaths Report** - Monthly death listings
10. **COD Summary** - Cause of death summary across all months
11. **Statement of Inpatient** - Yearly summary report
12. **Non-Insured Report** - Non-insured patient listing

### VBA Forms

- **Daily Bed Entry** - Enter daily statistics for all wards
- **Record Admission** - Individual patient admission details
- **Record Death** - Individual death records with cause of death
- **Record Ages Group** - Bulk age group entry mode
- **Refresh Reports** - Update death and COD summary sheets

## 🔧 Configuration

### Ward Configuration (`wards_config.json`)

Customize your hospital's wards by editing this file. Example:

```json
{
  "code": "ICU",
  "name": "Intensive Care Unit",
  "bed_complement": 12,
  "is_emergency": true,
  "display_order": 10
}
```

**Quick steps to add a ward:**
1. Open `wards_config.json`
2. Copy an existing ward entry
3. Modify the values
4. Save and rebuild the workbook

See [QUICK_START_WARDS.md](QUICK_START_WARDS.md) for detailed instructions.

### Carry Forward Data

At year-end, export remaining patients:
```bash
python build_workbook.py --year 2026
# Use VBA "Export Year-End" button in Excel
```

Start next year with previous year's data:
```bash
python build_workbook.py --year 2027 --carry-forward carry_forward_2026.json
```

## 📊 Key Performance Indicators (KPIs)

The system automatically calculates:
- Average Daily Bed Occupancy
- Average Length of Stay
- Bed Turnover Interval
- Bed Turnover Rate
- Percentage of Occupancy
- Death Rate

## 🛠️ Building the Workbook

### Basic Build
```bash
python build_workbook.py --year 2026
```

### Advanced Options
```bash
# Specify output directory
python build_workbook.py --year 2026 --output-dir /path/to/output

# Carry forward from previous year
python build_workbook.py --year 2026 --carry-forward carry_forward_2025.json

# Skip VBA injection (creates .xlsx instead of .xlsm)
python build_workbook.py --year 2026 --skip-vba
```

## 📁 Project Structure

```
bedutilization/
├── build_workbook.py          # Main build script
├── config.py                  # Configuration & ward definitions
├── phase1_structure.py        # Excel structure builder (openpyxl)
├── phase2_vba.py             # VBA macro injector (win32com)
├── wards_config.json         # ⭐ Ward configuration (user-editable)
├── carry_forward_2026.json   # Year-end rollover data
├── README.md                 # This file
├── WARD_CONFIGURATION_GUIDE.md
├── QUICK_START_WARDS.md
├── CHANGELOG.md
└── ocr_tool/                 # OCR tool for handwritten forms (in development)
```

## 🔐 Excel Macro Security

To use the VBA features:
1. Open Excel Options → Trust Center → Trust Center Settings
2. Enable "Trust access to the VBA project object model"
3. Add the project folder to Trusted Locations (recommended)

## 🤝 Support & Maintenance

### Common Tasks

**Add a new ward:**
- Edit `wards_config.json`
- Rebuild workbook

**Update bed complement:**
- Edit ward in `wards_config.json`
- Rebuild workbook

**Troubleshooting:**
- Check console output during build
- Verify JSON syntax at https://jsonlint.com
- Ensure all required fields are present

### Error Messages

| Error | Solution |
|-------|----------|
| "No wards found" | Check `wards_config.json` has a "wards" array |
| "Missing required field" | Ensure all fields present: code, name, bed_complement, is_emergency, display_order |
| "Invalid JSON" | Validate JSON syntax, check for missing commas |

## 📝 License

Ghana Health Service - Hohoe Municipal Hospital

## 🙏 Credits

Developed for the Ghana Health Service to streamline hospital bed utilization tracking and reporting.

---

**Last Updated:** 2026-02-07
**Version:** 2.0 (Ward Configuration Management Update)
