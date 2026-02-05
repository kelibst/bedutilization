"""
Phase 1: Build workbook structure using openpyxl
Creates all sheets, Excel Tables, formatting, formulas, and data validation.
"""
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from config import WorkbookConfig

# ── Style constants ──────────────────────────────────────────────────────────

HEADER_FONT = Font(name="Calibri", bold=True, size=14)
SUBHEADER_FONT = Font(name="Calibri", bold=True, size=12)
LABEL_FONT = Font(name="Calibri", bold=True, size=10)
NORMAL_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", bold=True, size=10)

CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT_WHITE = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
LIGHT_BLUE_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
LIGHT_GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
LIGHT_YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GRAY_FILL = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")

TABLE_STYLE = TableStyleInfo(
    name="TableStyleLight9",
    showFirstColumn=False, showLastColumn=False,
    showRowStripes=True, showColumnStripes=False,
)


def _apply_border_range(ws, min_row, max_row, min_col, max_col):
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).border = THIN_BORDER


# ═══════════════════════════════════════════════════════════════════════════════
# CONTROL SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def build_control_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.active
    ws.title = "Control"
    ws.sheet_properties.tabColor = "1F4E79"

    # Header
    ws.merge_cells("A1:H1")
    c = ws.cell(row=1, column=1, value="GHANA HEALTH SERVICE")
    c.font = HEADER_FONT
    c.alignment = CENTER

    ws.merge_cells("A2:H2")
    c = ws.cell(row=2, column=1, value=f"BED UTILIZATION MANAGEMENT SYSTEM - {config.year}")
    c.font = SUBHEADER_FONT
    c.alignment = CENTER

    ws.merge_cells("A3:H3")
    c = ws.cell(row=3, column=1, value=config.hospital_name)
    c.font = Font(name="Calibri", bold=True, size=11)
    c.alignment = CENTER

    # Named cells for Year and Hospital
    ws.cell(row=5, column=1, value="Year:").font = BOLD_FONT
    ws.cell(row=5, column=2, value=config.year).font = BOLD_FONT

    ws.cell(row=5, column=4, value="Hospital:").font = BOLD_FONT
    ws.cell(row=5, column=5, value=config.hospital_name).font = NORMAL_FONT

    # Instructions
    ws.merge_cells("A7:H7")
    ws.cell(row=7, column=1, value="Use the buttons below to enter data and view reports.").font = NORMAL_FONT

    # Button placeholders (actual buttons added by phase2 via win32com)
    buttons = [
        (9,  "[ DAILY BED ENTRY ]",       "Enter daily admissions, discharges, deaths, transfers for each ward"),
        (11, "[ RECORD ADMISSION ]",       "Record individual patient admission details (for age/gender reports)"),
        (13, "[ RECORD DEATH ]",           "Record individual death details (for deaths report)"),
        (15, "[ RECORD AGES GROUP ]",      "Quick entry for age group admissions (bulk mode)"),
        (17, "[ REFRESH REPORTS ]",        "Update Deaths Report and COD Summary sheets"),
    ]
    for row, label, desc in buttons:
        ws.merge_cells(f"A{row}:C{row}")
        c = ws.cell(row=row, column=1, value=label)
        c.font = Font(name="Calibri", bold=True, size=11, color="1F4E79")
        c.alignment = CENTER
        c.fill = LIGHT_BLUE_FILL
        c.border = THIN_BORDER
        ws.cell(row=row, column=4, value=desc).font = NORMAL_FONT

    # ── Ward Configuration Table ─────────────────────────────────────────
    config_start = 18
    ws.cell(row=config_start, column=1, value="WARD CONFIGURATION").font = SUBHEADER_FONT

    headers = ["WardCode", "WardName", "BedComplement", "PrevYearRemaining", "IsEmergency", "DisplayOrder"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=config_start + 1, column=col, value=h)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER

    for i, ward in enumerate(config.WARDS):
        r = config_start + 2 + i
        vals = [ward.code, ward.name, ward.bed_complement,
                ward.prev_year_remaining, ward.is_emergency, ward.display_order]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = NORMAL_FONT
            c.alignment = CENTER
            c.border = THIN_BORDER

    # Create the table
    end_row = config_start + 1 + len(config.WARDS)
    tab_ref = f"A{config_start + 1}:F{end_row}"
    tbl = Table(displayName="tblWardConfig", ref=tab_ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    # ── Month Lookup Table ───────────────────────────────────────────────
    month_start = end_row + 2
    ws.cell(row=month_start, column=1, value="MONTH LOOKUP").font = SUBHEADER_FONT

    month_headers = ["MonthNum", "MonthName", "DaysInMonth"]
    for col, h in enumerate(month_headers, 1):
        c = ws.cell(row=month_start + 1, column=col, value=h)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER

    for m in range(1, 13):
        r = month_start + 1 + m
        ws.cell(row=r, column=1, value=m).font = NORMAL_FONT
        ws.cell(row=r, column=2, value=config.MONTH_NAMES[m - 1]).font = NORMAL_FONT
        ws.cell(row=r, column=3, value=config.days_in_month(m)).font = NORMAL_FONT
        for col in range(1, 4):
            ws.cell(row=r, column=col).alignment = CENTER
            ws.cell(row=r, column=col).border = THIN_BORDER

    month_end = month_start + 13
    tbl2 = Table(displayName="tblMonthDays", ref=f"A{month_start + 1}:C{month_end}")
    tbl2.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl2)

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14


# ═══════════════════════════════════════════════════════════════════════════════
# DATA SHEETS (hidden tables)
# ═══════════════════════════════════════════════════════════════════════════════

def build_daily_data_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("DailyData")
    ws.sheet_properties.tabColor = "808080"

    headers = [
        "EntryDate", "Month", "WardCode", "Admissions", "Discharges",
        "Deaths", "DeathsUnder24Hrs", "TransfersIn", "TransfersOut",
        "PrevRemaining", "Remaining", "EntryTimestamp"
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    # Seed row (VBA calculates PrevRemaining and Remaining as VALUES)
    ws.cell(row=2, column=1, value="")  # EntryDate (VBA fills)
    ws.cell(row=2, column=2, value="")  # Month (VBA fills)
    ws.cell(row=2, column=3, value="")  # WardCode (VBA fills)
    ws.cell(row=2, column=4, value="")  # Admissions (VBA fills)
    ws.cell(row=2, column=5, value="")  # Discharges (VBA fills)
    ws.cell(row=2, column=6, value="")  # Deaths (VBA fills)
    ws.cell(row=2, column=7, value="")  # DeathsUnder24Hrs (VBA fills)
    ws.cell(row=2, column=8, value="")  # TransfersIn (VBA fills)
    ws.cell(row=2, column=9, value="")  # TransfersOut (VBA fills)
    ws.cell(row=2, column=10, value="")  # PrevRemaining (VBA calculates)
    ws.cell(row=2, column=11, value="")  # Remaining (VBA calculates)
    ws.cell(row=2, column=12, value="")  # EntryTimestamp (VBA fills)

    tbl = Table(displayName="tblDaily", ref="A1:L2")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 12
    for c in "DEFGHIJ":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["K"].width = 12
    ws.column_dimensions["L"].width = 18


def build_admissions_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Admissions")
    ws.sheet_properties.tabColor = "808080"

    headers = [
        "AdmissionID", "AdmissionDate", "Month", "WardCode", "PatientID",
        "PatientName", "Age", "AgeUnit", "Sex", "NHIS", "EntryTimestamp"
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    ws.cell(row=2, column=1, value="")

    tbl = Table(displayName="tblAdmissions", ref="A1:K2")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 25
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 6
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 18


def build_deaths_data_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("DeathsData")
    ws.sheet_properties.tabColor = "808080"

    headers = [
        "DeathID", "DateOfDeath", "Month", "WardCode", "FolderNumber",
        "NameOfDeceased", "Age", "AgeUnit", "Sex", "NHIS",
        "CauseOfDeath", "DeathWithin24Hrs", "EntryTimestamp"
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    ws.cell(row=2, column=1, value="")

    tbl = Table(displayName="tblDeaths", ref="A1:M2")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 25
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 6
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 25
    ws.column_dimensions["L"].width = 16
    ws.column_dimensions["M"].width = 18


def build_transfers_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("TransfersData")
    ws.sheet_properties.tabColor = "808080"

    headers = [
        "TransferID", "TransferDate", "Month", "FromWardCode",
        "ToWardCode", "PatientID", "PatientName", "EntryTimestamp"
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    ws.cell(row=2, column=1, value="")

    tbl = Table(displayName="tblTransfers", ref="A1:H2")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)


# ═══════════════════════════════════════════════════════════════════════════════
# WARD REPORT SHEETS (9 sheets, one per ward)
# ═══════════════════════════════════════════════════════════════════════════════

def build_ward_sheet(wb: Workbook, config: WorkbookConfig, ward):
    ws = wb.create_sheet(ward.name)

    # Tab colors: wards get distinct colors
    ward_colors = {
        "MW": "4472C4", "FW": "ED7D31", "CW": "A5A5A5",
        "BF": "FFC000", "BG": "5B9BD5", "BH": "70AD47",
        "NICU": "7030A0", "MAE": "FF0000", "FAE": "FF69B4",
    }
    ws.sheet_properties.tabColor = ward_colors.get(ward.code, "000000")

    current_row = 1
    for month_num in range(1, 13):
        month_name = config.MONTH_NAMES[month_num - 1]
        days = config.days_in_month(month_num)

        # ── Header block ─────────────────────────────────────────────
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=8)
        c = ws.cell(row=current_row, column=1, value="GHANA HEALTH SERVICE")
        c.font = HEADER_FONT
        c.alignment = CENTER
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=8)
        c = ws.cell(row=current_row, column=1, value="DAILY BED UTILIZATION FORM")
        c.font = SUBHEADER_FONT
        c.alignment = CENTER
        current_row += 1

        # Hospital / Ward / Month row
        ws.cell(row=current_row, column=1, value="Hospital:").font = BOLD_FONT
        ws.cell(row=current_row, column=2, value=config.hospital_name).font = NORMAL_FONT
        ws.cell(row=current_row, column=4, value="Ward:").font = BOLD_FONT
        ws.cell(row=current_row, column=5, value=ward.name.upper()).font = NORMAL_FONT
        ws.cell(row=current_row, column=6, value="MONTH").font = BOLD_FONT
        ws.cell(row=current_row, column=7, value=month_name).font = NORMAL_FONT
        current_row += 1

        # Previous remaining
        ws.cell(row=current_row, column=1, value="Number of patients remaining as at last day of previous month").font = LABEL_FONT
        if month_num == 1:
            formula = (
                f'=IFERROR(INDEX(tblWardConfig[PrevYearRemaining],'
                f'MATCH("{ward.code}",tblWardConfig[WardCode],0)),0)'
            )
        else:
            prev_month = month_num - 1
            prev_last_day = config.days_in_month(prev_month)
            formula = (
                f'=IFERROR(SUMIFS(tblDaily[Remaining],'
                f'tblDaily[EntryDate],DATE({config.year},{prev_month},{prev_last_day}),'
                f'tblDaily[WardCode],"{ward.code}"),0)'
            )
        c = ws.cell(row=current_row, column=8, value=formula)
        c.font = BOLD_FONT
        c.fill = LIGHT_YELLOW_FILL
        c.border = THIN_BORDER
        prev_remaining_row = current_row
        current_row += 1

        # Bed complement
        ws.cell(row=current_row, column=1, value="Bed complement").font = LABEL_FONT
        bed_formula = (
            f'=IFERROR(INDEX(tblWardConfig[BedComplement],'
            f'MATCH("{ward.code}",tblWardConfig[WardCode],0)),0)'
        )
        c = ws.cell(row=current_row, column=8, value=bed_formula)
        c.font = BOLD_FONT
        c.fill = LIGHT_GREEN_FILL
        c.border = THIN_BORDER
        current_row += 1

        # ── Column headers ───────────────────────────────────────────
        col_headers = [
            "Day of\nthe\nMonth", "Admissions", "Discharges", "Deaths",
            "Deaths\n<24Hrs", "Transfers-\nIn", "Transfers-\nOut",
            "No. of Patients\nRemaining In\nWard", "Malaria\nCases"
        ]
        for col, header in enumerate(col_headers, 1):
            c = ws.cell(row=current_row, column=col, value=header)
            c.font = HEADER_FONT_WHITE
            c.fill = HEADER_FILL
            c.alignment = CENTER_WRAP
            c.border = THIN_BORDER
        header_row = current_row
        current_row += 1

        # ── Daily rows (1-31) ────────────────────────────────────────
        data_start = current_row
        for day in range(1, 32):
            ws.cell(row=current_row, column=1, value=day).font = NORMAL_FONT
            ws.cell(row=current_row, column=1).alignment = CENTER
            ws.cell(row=current_row, column=1).border = THIN_BORDER

            if day <= days:
                date_ref = f"DATE({config.year},{month_num},{day})"
                fields = ["Admissions", "Discharges", "Deaths",
                           "DeathsUnder24Hrs", "TransfersIn", "TransfersOut", "Remaining", "MalariaCases"]
                for col_idx, field_name in enumerate(fields, 2):
                    formula = (
                        f'=IFERROR(IF(SUMIFS(tblDaily[{field_name}],'
                        f'tblDaily[EntryDate],{date_ref},'
                        f'tblDaily[WardCode],"{ward.code}")=0,"",'
                        f'SUMIFS(tblDaily[{field_name}],'
                        f'tblDaily[EntryDate],{date_ref},'
                        f'tblDaily[WardCode],"{ward.code}")),"")'
                    )
                    c = ws.cell(row=current_row, column=col_idx, value=formula)
                    c.font = NORMAL_FONT
                    c.alignment = CENTER
                    c.border = THIN_BORDER
            else:
                for col_idx in range(2, 10):
                    c = ws.cell(row=current_row, column=col_idx)
                    c.fill = GRAY_FILL
                    c.border = THIN_BORDER

            current_row += 1

        # ── TOTAL row ────────────────────────────────────────────────
        c = ws.cell(row=current_row, column=1, value="TOTAL")
        c.font = BOLD_FONT
        c.alignment = CENTER
        c.fill = TOTAL_FILL
        c.border = THIN_BORDER

        total_fields = ["Admissions", "Discharges", "Deaths",
                        "DeathsUnder24Hrs", "TransfersIn", "TransfersOut", "MalariaCases"]
        for col_idx, field_name in enumerate(total_fields, 2):
            formula = (
                f'=SUMIFS(tblDaily[{field_name}],'
                f'tblDaily[Month],{month_num},'
                f'tblDaily[WardCode],"{ward.code}")'
            )
            c = ws.cell(row=current_row, column=col_idx, value=formula)
            c.font = BOLD_FONT
            c.alignment = CENTER
            c.fill = TOTAL_FILL
            c.border = THIN_BORDER

        # Patient Days (sum of daily remaining) - column 8
        pd_formula = (
            f'=SUMIFS(tblDaily[Remaining],'
            f'tblDaily[Month],{month_num},'
            f'tblDaily[WardCode],"{ward.code}")'
        )
        c = ws.cell(row=current_row, column=8, value=pd_formula)
        c.font = BOLD_FONT
        c.alignment = CENTER
        c.fill = TOTAL_FILL
        c.border = THIN_BORDER

        current_row += 2  # spacer

    # Column widths
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["I"].width = 12  # Malaria Cases

    # Print setup
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToPage = True


# ═══════════════════════════════════════════════════════════════════════════════
# MONTHLY SUMMARY SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def build_monthly_summary_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Monthly Summary")
    ws.sheet_properties.tabColor = "1F4E79"

    current_row = 1
    for month_num in range(1, 13):
        month_name = config.MONTH_NAMES[month_num - 1]
        days = config.days_in_month(month_num)

        # Header
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=16)
        c = ws.cell(row=current_row, column=1, value="GHANA HEALTH SERVICE")
        c.font = HEADER_FONT
        c.alignment = CENTER
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=16)
        c = ws.cell(row=current_row, column=1,
                    value=f"MONTHLY BED UTILIZATION FORM - {month_name}, {config.year}")
        c.font = SUBHEADER_FONT
        c.alignment = CENTER
        current_row += 1

        # Column headers
        col_headers = [
            "WARD", "Patients on bed\nat the beginning\nof the month",
            "Bed Complement", "Admissions", "Discharges", "Deaths",
            "Deaths\n<24Hrs", "Patient\nDays", "Transfer In", "Transfer\nOut",
            "Average\nDaily Bed\nOccupancy", "Average\nLength\nof Stay",
            "Bed\nTurnover\nInterval", "Bed\nTurnover\nRate",
            "Percentage\nof\nOccupancy", "Death\nRate"
        ]
        for col, header in enumerate(col_headers, 1):
            c = ws.cell(row=current_row, column=col, value=header)
            c.font = HEADER_FONT_WHITE
            c.fill = HEADER_FILL
            c.alignment = CENTER_WRAP
            c.border = THIN_BORDER
        header_row = current_row
        current_row += 1

        # ── Ward rows ────────────────────────────────────────────────
        for ward in config.WARDS:
            r = current_row
            wc = ward.code

            # Col A: Ward name
            ws.cell(row=r, column=1, value=ward.name).font = BOLD_FONT
            ws.cell(row=r, column=1).border = THIN_BORDER

            # Col B: Patients at beginning of month
            if month_num == 1:
                f_beg = (
                    f'=IFERROR(INDEX(tblWardConfig[PrevYearRemaining],'
                    f'MATCH("{wc}",tblWardConfig[WardCode],0)),0)'
                )
            else:
                pm = month_num - 1
                pld = config.days_in_month(pm)
                f_beg = (
                    f'=IFERROR(SUMIFS(tblDaily[Remaining],'
                    f'tblDaily[EntryDate],DATE({config.year},{pm},{pld}),'
                    f'tblDaily[WardCode],"{wc}"),0)'
                )
            ws.cell(row=r, column=2, value=f_beg).font = NORMAL_FONT

            # Col C: Bed Complement
            f_bc = f'=IFERROR(INDEX(tblWardConfig[BedComplement],MATCH("{wc}",tblWardConfig[WardCode],0)),0)'
            ws.cell(row=r, column=3, value=f_bc).font = NORMAL_FONT

            # Col D-F,G: Admissions, Discharges, Deaths, Deaths<24Hrs
            daily_fields = {
                4: "Admissions", 5: "Discharges", 6: "Deaths", 7: "DeathsUnder24Hrs"
            }
            for col_num, field in daily_fields.items():
                f = f'=SUMIFS(tblDaily[{field}],tblDaily[Month],{month_num},tblDaily[WardCode],"{wc}")'
                ws.cell(row=r, column=col_num, value=f).font = NORMAL_FONT

            # Col H: Patient Days
            f_pd = f'=SUMIFS(tblDaily[Remaining],tblDaily[Month],{month_num},tblDaily[WardCode],"{wc}")'
            ws.cell(row=r, column=8, value=f_pd).font = NORMAL_FONT

            # Col I-J: Transfers In, Transfers Out
            f_ti = f'=SUMIFS(tblDaily[TransfersIn],tblDaily[Month],{month_num},tblDaily[WardCode],"{wc}")'
            f_to = f'=SUMIFS(tblDaily[TransfersOut],tblDaily[Month],{month_num},tblDaily[WardCode],"{wc}")'
            ws.cell(row=r, column=9, value=f_ti).font = NORMAL_FONT
            ws.cell(row=r, column=10, value=f_to).font = NORMAL_FONT

            # Col K: Average Daily Bed Occupancy = Patient Days / Days
            bc = get_column_letter(8)
            ws.cell(row=r, column=11, value=f'=IFERROR({bc}{r}/{days},0)').font = NORMAL_FONT

            # Col L: Average Length of Stay = Patient Days / (Discharges + Deaths)
            ws.cell(row=r, column=12, value=f'=IFERROR({bc}{r}/(E{r}+F{r}),0)').font = NORMAL_FONT

            # Col M: Bed Turnover Interval = (BC*Days - PD) / (Disch + Deaths)
            ws.cell(row=r, column=13, value=f'=IFERROR((C{r}*{days}-H{r})/(E{r}+F{r}),0)').font = NORMAL_FONT

            # Col N: Bed Turnover Rate = (Disch + Deaths) / BC
            ws.cell(row=r, column=14, value=f'=IFERROR((E{r}+F{r})/C{r},0)').font = NORMAL_FONT

            # Col O: % Occupancy = (PD / (BC * Days)) * 100
            ws.cell(row=r, column=15, value=f'=IFERROR((H{r}/(C{r}*{days}))*100,0)').font = NORMAL_FONT

            # Col P: Death Rate = Deaths / (Admissions + Patients at beginning) * 100
            ws.cell(row=r, column=16, value=f'=IFERROR((F{r}/(D{r}+B{r}))*100,0)').font = NORMAL_FONT

            # Format numbers
            for col in range(11, 17):
                ws.cell(row=r, column=col).number_format = '0.00'

            # Borders
            for col in range(1, 17):
                ws.cell(row=r, column=col).alignment = CENTER
                ws.cell(row=r, column=col).border = THIN_BORDER

            current_row += 1

        # ── TOTAL row ────────────────────────────────────────────────
        r = current_row
        ws.cell(row=r, column=1, value="TOTAL").font = BOLD_FONT
        ws.cell(row=r, column=1).fill = TOTAL_FILL

        # Sum columns B through J across the ward rows
        first_ward_row = r - len(config.WARDS)
        last_ward_row = r - 1
        for col in range(2, 11):
            cl = get_column_letter(col)
            ws.cell(row=r, column=col,
                    value=f'=SUM({cl}{first_ward_row}:{cl}{last_ward_row})').font = BOLD_FONT

        # KPI totals use total values
        ws.cell(row=r, column=11, value=f'=IFERROR(H{r}/{days},0)').font = BOLD_FONT
        ws.cell(row=r, column=12, value=f'=IFERROR(H{r}/(E{r}+F{r}),0)').font = BOLD_FONT
        ws.cell(row=r, column=13, value=f'=IFERROR((C{r}*{days}-H{r})/(E{r}+F{r}),0)').font = BOLD_FONT
        ws.cell(row=r, column=14, value=f'=IFERROR((E{r}+F{r})/C{r},0)').font = BOLD_FONT
        ws.cell(row=r, column=15, value=f'=IFERROR((H{r}/(C{r}*{days}))*100,0)').font = BOLD_FONT
        ws.cell(row=r, column=16, value=f'=IFERROR((F{r}/(D{r}+B{r}))*100,0)').font = BOLD_FONT

        for col in range(1, 17):
            ws.cell(row=r, column=col).alignment = CENTER
            ws.cell(row=r, column=col).border = THIN_BORDER
            ws.cell(row=r, column=col).fill = TOTAL_FILL
        for col in range(11, 17):
            ws.cell(row=r, column=col).number_format = '0.00'
        current_row += 1

        # ── Emergency subtotal row ───────────────────────────────────
        r = current_row
        ws.cell(row=r, column=1, value="Emergency").font = BOLD_FONT
        ws.cell(row=r, column=1).fill = LIGHT_YELLOW_FILL

        emer_wards = [w for w in config.WARDS if w.is_emergency]
        for col in range(2, 11):
            parts = []
            for ew in emer_wards:
                # Find the row for this emergency ward
                ew_idx = config.WARDS.index(ew)
                ew_row = first_ward_row + ew_idx
                parts.append(f'{get_column_letter(col)}{ew_row}')
            ws.cell(row=r, column=col, value=f'={"+".join(parts)}').font = BOLD_FONT

        ws.cell(row=r, column=11, value=f'=IFERROR(H{r}/{days},0)').font = BOLD_FONT
        ws.cell(row=r, column=12, value=f'=IFERROR(H{r}/(E{r}+F{r}),0)').font = BOLD_FONT
        ws.cell(row=r, column=13, value=f'=IFERROR((C{r}*{days}-H{r})/(E{r}+F{r}),0)').font = BOLD_FONT
        ws.cell(row=r, column=14, value=f'=IFERROR((E{r}+F{r})/C{r},0)').font = BOLD_FONT
        ws.cell(row=r, column=15, value=f'=IFERROR((H{r}/(C{r}*{days}))*100,0)').font = BOLD_FONT
        ws.cell(row=r, column=16, value=f'=IFERROR((F{r}/(D{r}+B{r}))*100,0)').font = BOLD_FONT

        for col in range(1, 17):
            ws.cell(row=r, column=col).alignment = CENTER
            ws.cell(row=r, column=col).border = THIN_BORDER
            ws.cell(row=r, column=col).fill = LIGHT_YELLOW_FILL
        for col in range(11, 17):
            ws.cell(row=r, column=col).number_format = '0.00'

        current_row += 2  # spacer before next month

    # Column widths
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    for i in range(4, 11):
        ws.column_dimensions[get_column_letter(i)].width = 12
    for i in range(11, 17):
        ws.column_dimensions[get_column_letter(i)].width = 12

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToPage = True


# ═══════════════════════════════════════════════════════════════════════════════
# AGES SUMMARY SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def build_ages_summary_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Ages Summary")
    ws.sheet_properties.tabColor = "7030A0"

    # Layout: 12 horizontal sections (one per month), each occupying ~7 columns
    # For each month section:
    # Row 1: Month name
    # Row 2: Category headers (both ins and non ins | NON INS)
    # Row 3: Male | Female | Male | Female subheaders
    # Row 4-15: Age group data
    # Row 16: Total

    SECTION_WIDTH = 7  # columns per month section (age_label + 6 data cols)

    # Row 1: Month headers
    for m in range(1, 13):
        start_col = 1 + (m - 1) * SECTION_WIDTH
        ws.merge_cells(start_row=1, start_column=start_col,
                       end_row=1, end_column=start_col + SECTION_WIDTH - 1)
        c = ws.cell(row=1, column=start_col, value=config.MONTH_NAMES[m - 1])
        c.font = SUBHEADER_FONT
        c.alignment = CENTER
        c.fill = HEADER_FILL
        c.font = HEADER_FONT_WHITE

    # Row 2: Category headers
    for m in range(1, 13):
        sc = 1 + (m - 1) * SECTION_WIDTH
        # Col 0: blank (age group label)
        # Cols 1-2: "both ins and non ins" Male/Female
        ws.merge_cells(start_row=2, start_column=sc + 1, end_row=2, end_column=sc + 2)
        c = ws.cell(row=2, column=sc + 1, value="both ins and non ins")
        c.font = LABEL_FONT
        c.alignment = CENTER
        c.fill = LIGHT_BLUE_FILL

        # Cols 3-4: "NON INS" Male/Female
        ws.merge_cells(start_row=2, start_column=sc + 3, end_row=2, end_column=sc + 4)
        c = ws.cell(row=2, column=sc + 3, value="NON INS")
        c.font = LABEL_FONT
        c.alignment = CENTER
        c.fill = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")

        # Cols 5-6: "INS" (derived = total - non-ins, but we use COUNTIFS)
        # Actually from the screenshot, INS is shown separately
        # Let me add INS columns too
        ws.merge_cells(start_row=2, start_column=sc + 5, end_row=2, end_column=sc + 6)
        c = ws.cell(row=2, column=sc + 5, value="INS")
        c.font = LABEL_FONT
        c.alignment = CENTER
        c.fill = LIGHT_GREEN_FILL

    # Row 3: Male/Female subheaders
    for m in range(1, 13):
        sc = 1 + (m - 1) * SECTION_WIDTH
        ws.cell(row=3, column=sc, value="Age Group").font = BOLD_FONT
        ws.cell(row=3, column=sc).alignment = CENTER
        ws.cell(row=3, column=sc).border = THIN_BORDER
        for offset, label in [(1, "Male"), (2, "FEMALE"),
                               (3, "MALE"), (4, "FRMALE"),
                               (5, "MALE"), (6, "FRMALE")]:
            c = ws.cell(row=3, column=sc + offset, value=label)
            c.font = BOLD_FONT
            c.alignment = CENTER
            c.border = THIN_BORDER

    # Rows 4-15: Age group data with COUNTIFS formulas
    for m in range(1, 13):
        sc = 1 + (m - 1) * SECTION_WIDTH

        for ag_idx, (ag_label, age_unit, age_min, age_max) in enumerate(config.AGE_GROUPS):
            r = 4 + ag_idx
            ws.cell(row=r, column=sc, value=ag_label).font = NORMAL_FONT
            ws.cell(row=r, column=sc).alignment = CENTER
            ws.cell(row=r, column=sc).border = THIN_BORDER

            # Build COUNTIFS for each combination
            # Both ins and non-ins Male (col sc+1)
            base = (
                f'COUNTIFS(tblAdmissions[Month],{m},'
                f'tblAdmissions[AgeUnit],"{age_unit}",'
                f'tblAdmissions[Age],">="&{age_min},'
                f'tblAdmissions[Age],"<="&{age_max},'
            )
            # Total Male
            ws.cell(row=r, column=sc + 1,
                    value=f'={base}tblAdmissions[Sex],"M")').font = NORMAL_FONT
            # Total Female
            ws.cell(row=r, column=sc + 2,
                    value=f'={base}tblAdmissions[Sex],"F")').font = NORMAL_FONT
            # Non-Insured Male
            ws.cell(row=r, column=sc + 3,
                    value=f'={base}tblAdmissions[Sex],"M",tblAdmissions[NHIS],"Non-Insured")').font = NORMAL_FONT
            # Non-Insured Female
            ws.cell(row=r, column=sc + 4,
                    value=f'={base}tblAdmissions[Sex],"F",tblAdmissions[NHIS],"Non-Insured")').font = NORMAL_FONT
            # Insured Male
            ws.cell(row=r, column=sc + 5,
                    value=f'={base}tblAdmissions[Sex],"M",tblAdmissions[NHIS],"Insured")').font = NORMAL_FONT
            # Insured Female
            ws.cell(row=r, column=sc + 6,
                    value=f'={base}tblAdmissions[Sex],"F",tblAdmissions[NHIS],"Insured")').font = NORMAL_FONT

            for col_off in range(1, 7):
                ws.cell(row=r, column=sc + col_off).alignment = CENTER
                ws.cell(row=r, column=sc + col_off).border = THIN_BORDER

        # Total row
        total_row = 4 + len(config.AGE_GROUPS)
        ws.cell(row=total_row, column=sc, value="Total").font = BOLD_FONT
        ws.cell(row=total_row, column=sc).alignment = CENTER
        ws.cell(row=total_row, column=sc).fill = TOTAL_FILL
        ws.cell(row=total_row, column=sc).border = THIN_BORDER
        for col_off in range(1, 7):
            cl = get_column_letter(sc + col_off)
            ws.cell(row=total_row, column=sc + col_off,
                    value=f'=SUM({cl}4:{cl}{total_row - 1})').font = BOLD_FONT
            ws.cell(row=total_row, column=sc + col_off).alignment = CENTER
            ws.cell(row=total_row, column=sc + col_off).fill = TOTAL_FILL
            ws.cell(row=total_row, column=sc + col_off).border = THIN_BORDER

    # Set column widths
    for m in range(1, 13):
        sc = 1 + (m - 1) * SECTION_WIDTH
        ws.column_dimensions[get_column_letter(sc)].width = 8
        for off in range(1, 7):
            ws.column_dimensions[get_column_letter(sc + off)].width = 10


# ═══════════════════════════════════════════════════════════════════════════════
# DEATHS REPORT SHEET (VBA-refreshed, but we set up the structure)
# ═══════════════════════════════════════════════════════════════════════════════

def build_deaths_report_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Deaths Report")
    ws.sheet_properties.tabColor = "C00000"

    # Create 12 month sections with headers
    current_row = 1
    for month_num in range(1, 13):
        month_name = config.MONTH_NAMES[month_num - 1]

        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=8)
        c = ws.cell(row=current_row, column=1,
                    value=f"DEATHS FOR {month_name}")
        c.font = SUBHEADER_FONT
        c.alignment = CENTER
        current_row += 1

        headers = ["S/N", "Folder Number", "Date of Death",
                   "Name of Deceased", "Age", "Sex", "Ward", "NHIS"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=current_row, column=col, value=h)
            c.font = HEADER_FONT_WHITE
            c.fill = HEADER_FILL
            c.alignment = CENTER
            c.border = THIN_BORDER
        current_row += 1

        # Leave space for data (VBA will populate this)
        # Reserve 40 rows per month
        current_row += 40

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 6
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 10


# ═══════════════════════════════════════════════════════════════════════════════
# COD SUMMARY SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def build_cod_summary_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("COD Summary")
    ws.sheet_properties.tabColor = "C00000"

    ws.merge_cells("A1:N1")
    c = ws.cell(row=1, column=1, value="CAUSE OF DEATH SUMMARY")
    c.font = HEADER_FONT
    c.alignment = CENTER

    # Headers: Cause of Death | Jan | Feb | ... | Dec | Total
    headers = ["Cause of Death"] + list(config.MONTH_NAMES) + ["TOTAL"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER

    # VBA will populate rows 3+ with distinct causes and COUNTIFS
    # Leave as placeholder
    ws.cell(row=3, column=1, value="(Click 'Refresh Reports' to populate)").font = NORMAL_FONT

    ws.column_dimensions["A"].width = 30
    for col in range(2, 15):
        ws.column_dimensions[get_column_letter(col)].width = 10


# ═══════════════════════════════════════════════════════════════════════════════
# STATEMENT OF INPATIENT SHEET (Yearly Summary)
# ═══════════════════════════════════════════════════════════════════════════════

def build_statement_of_inpatient_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Statement of Inpatient")
    ws.sheet_properties.tabColor = "0000FF"

    ws.merge_cells("A1:P1")
    c = ws.cell(row=1, column=1, value="GHANA HEALTH SERVICE - STATEMENT OF INPATIENT (YEARLY SUMMARY)")
    c.font = HEADER_FONT
    c.alignment = CENTER

    # Headers same as Monthly Summary
    col_headers = [
        "WARD", "Patients on bed\nat start of year",
        "Bed Complement", "Admissions", "Discharges", "Deaths",
        "Deaths\n<24Hrs", "Patient\nDays", "Transfer In", "Transfer\nOut",
        "Average\nDaily Bed\nOccupancy", "Average\nLength\nof Stay",
        "Bed\nTurnover\nInterval", "Bed\nTurnover\nRate",
        "Percentage\nof\nOccupancy", "Death\nRate"
    ]
    for col, header in enumerate(col_headers, 1):
        c = ws.cell(row=3, column=col, value=header)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.alignment = CENTER_WRAP
        c.border = THIN_BORDER
    
    current_row = 4
    days_in_year = 365 + (1 if config.year % 4 == 0 else 0)

    for ward in config.WARDS:
        r = current_row
        wc = ward.code
        
        # A: Name
        ws.cell(row=r, column=1, value=ward.name).font = BOLD_FONT
        
        # B: Start of year
        f_beg = (
            f'=IFERROR(INDEX(tblWardConfig[PrevYearRemaining],'
            f'MATCH("{wc}",tblWardConfig[WardCode],0)),0)'
        )
        ws.cell(row=r, column=2, value=f_beg).font = NORMAL_FONT
        
        # C: BC
        f_bc = f'=IFERROR(INDEX(tblWardConfig[BedComplement],MATCH("{wc}",tblWardConfig[WardCode],0)),0)'
        ws.cell(row=r, column=3, value=f_bc).font = NORMAL_FONT
        
        # D-G, I-J: Sums
        ws.cell(row=r, column=4, value=f'=SUMIFS(tblDaily[Admissions],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=5, value=f'=SUMIFS(tblDaily[Discharges],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=6, value=f'=SUMIFS(tblDaily[Deaths],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=7, value=f'=SUMIFS(tblDaily[DeathsUnder24Hrs],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=8, value=f'=SUMIFS(tblDaily[Remaining],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=9, value=f'=SUMIFS(tblDaily[TransfersIn],tblDaily[WardCode],"{wc}")')
        ws.cell(row=r, column=10, value=f'=SUMIFS(tblDaily[TransfersOut],tblDaily[WardCode],"{wc}")')
        
        # KPIs
        bc_Ref = f"C{r}"
        pd_Ref = f"H{r}"
        adm_Ref = f"D{r}"
        dis_Ref = f"E{r}"
        dth_Ref = f"F{r}"
        beg_Ref = f"B{r}"
        
        # Av Occ = PD / 365
        ws.cell(row=r, column=11, value=f'=IFERROR({pd_Ref}/{days_in_year},0)')
        # Av LOS = PD / (Dis + Dth)
        ws.cell(row=r, column=12, value=f'=IFERROR({pd_Ref}/({dis_Ref}+{dth_Ref}),0)')
        # Turn Int = (BC*365 - PD) / (Dis + Dth)
        ws.cell(row=r, column=13, value=f'=IFERROR(({bc_Ref}*{days_in_year}-{pd_Ref})/({dis_Ref}+{dth_Ref}),0)')
        # Turn Rate = (Dis + Dth) / BC
        ws.cell(row=r, column=14, value=f'=IFERROR(({dis_Ref}+{dth_Ref})/{bc_Ref},0)')
        # % Occ = (PD / (BC*365)) * 100
        ws.cell(row=r, column=15, value=f'=IFERROR(({pd_Ref}/({bc_Ref}*{days_in_year}))*100,0)')
        # Death Rate
        ws.cell(row=r, column=16, value=f'=IFERROR(({dth_Ref}/({adm_Ref}+{beg_Ref}))*100,0)')
        
        current_row += 1

    # Format
    for r in range(4, current_row):
        for c in range(1, 17):
            ws.cell(row=r, column=c).border = THIN_BORDER
        for c in range(11, 17):
            ws.cell(row=r, column=c).number_format = '0.00'
            
    ws.column_dimensions["A"].width = 18
    for i in range(2, 17):
        ws.column_dimensions[get_column_letter(i)].width = 12


# ═══════════════════════════════════════════════════════════════════════════════
# NON-INSURED REPORT SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def build_non_insured_report_sheet(wb: Workbook, config: WorkbookConfig):
    ws = wb.create_sheet("Non-Insured Report")
    ws.sheet_properties.tabColor = "FFA500"

    ws.merge_cells("A1:J1")
    c = ws.cell(row=1, column=1, value="NON-INSURED ATTENDANTS REPORT")
    c.font = HEADER_FONT
    c.alignment = CENTER

    headers = ["S/N", "Date", "Month", "Ward", "Patient ID", "Name", "Age", "Sex", "Amount", "Status"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER

    ws.cell(row=3, column=1, value="(Click 'Refresh Reports' to populate)").font = NORMAL_FONT
    
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 25
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 6
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 12


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN BUILD FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def build_structure(config: WorkbookConfig, output_path: str):
    wb = Workbook()

    # 1. Control sheet (landing page)
    build_control_sheet(wb, config)

    # 2. Data sheets (hidden tables)
    build_daily_data_sheet(wb, config)
    build_admissions_sheet(wb, config)
    build_deaths_data_sheet(wb, config)
    build_transfers_sheet(wb, config)

    # 3. Ward report sheets (9 wards)
    for ward in config.WARDS:
        build_ward_sheet(wb, config, ward)

    # 4. Monthly Summary
    build_monthly_summary_sheet(wb, config)

    # 5. Ages Summary
    build_ages_summary_sheet(wb, config)

    # 6. Deaths Report
    build_deaths_report_sheet(wb, config)

    # 7. COD Summary
    build_cod_summary_sheet(wb, config)
    
    # 8. Statement of Inpatient
    build_statement_of_inpatient_sheet(wb, config)
    
    # 9. Non-Insured Report
    build_non_insured_report_sheet(wb, config)

    # Save
    wb.save(output_path)
    print(f"Phase 1 complete: {output_path}")
