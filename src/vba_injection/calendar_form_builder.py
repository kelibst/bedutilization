"""
Calendar Form Builder

Creates the frmCalendarPicker UserForm programmatically with all required controls.
This form provides a visual calendar interface for date selection, compatible with
64-bit Excel (replacement for MSComCtl2.DTPicker).
"""
from typing import Any
from .ui_helpers import add_label, add_combobox, add_button
from .utils import get_vba_path, read_vba_file


def create_calendar_picker_form(vbproj: Any) -> None:
    """
    Create the frmCalendarPicker UserForm with visual calendar interface.

    Layout:
    ┌──────────────────────────────────────┐
    │  [< Prev]  February 2026  [Next >]  │
    │  Month [▼] Year [▼]      [Today]    │
    ├──────────────────────────────────────┤
    │  Su  Mo  Tu  We  Th  Fr  Sa         │
    │                          1           │
    │   2   3   4   5   6   7   8         │
    │   9  10  11  12  13  14  15         │
    │  16  17  18  19  20  21  22         │
    │  23  24  25  26  27  28             │
    ├──────────────────────────────────────┤
    │         [Select]    [Cancel]         │
    └──────────────────────────────────────┘

    Args:
        vbproj: VBA project object to add the form to
    """
    # Create the UserForm
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm = 3
    form.Name = "frmCalendarPicker"

    # Set form properties
    form.Properties("Caption").Value = "Select Date"
    form.Properties("Width").Value = 325
    form.Properties("Height").Value = 300

    # Get form designer
    d = form.Designer

    # =========================================================================
    # Navigation controls (top row)
    # =========================================================================

    # Previous month button
    add_button(d, "btnPrev", "< Prev", 10, 10, 40, 22)

    # Month/Year display label (center)
    lblMonthYear = add_label(d, "lblMonthYear", "February 2026", 55, 12, 140, 18)
    lblMonthYear.Font.Size = 11
    lblMonthYear.Font.Bold = True
    lblMonthYear.TextAlign = 2  # fmTextAlignCenter

    # Next month button
    add_button(d, "btnNext", "Next >", 200, 10, 40, 22)

    # Select button (moved to top row, before Today)
    btnSelect = add_button(d, "btnSelect", "Select", 245, 10, 50, 22)
    btnSelect.Font.Size = 9
    btnSelect.BackColor = 0x90EE90  # Light green

    # =========================================================================
    # Month/Year selection controls (second row)
    # =========================================================================

    # Month label
    add_label(d, "lblMonthSelect", "Month:", 10, 40, 35, 18)

    # Month combobox
    add_combobox(d, "cmbMonth", 50, 38, 100, 20, style=2)  # DropDownList

    # Year label
    add_label(d, "lblYearSelect", "Year:", 160, 40, 30, 18)

    # Year combobox
    add_combobox(d, "cmbYear", 195, 38, 60, 20, style=2)  # DropDownList

    # Cancel button (second row, right side)
    btnCancel = add_button(d, "btnCancel", "Cancel", 260, 38, 50, 22)
    btnCancel.Font.Size = 9

    # Today button (second row, after cancel - or we can skip it to save space)
    # User can navigate using Prev/Next, so Today is optional

    # =========================================================================
    # Weekday headers (third row)
    # =========================================================================

    weekdays = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]
    for col, day_name in enumerate(weekdays):
        x = 10 + (col * 40)
        y = 68
        lbl = add_label(d, f"lblWeekday_{col}", day_name, x, y, 35, 18)
        lbl.Font.Bold = True
        lbl.TextAlign = 2  # fmTextAlignCenter

    # =========================================================================
    # Day grid (6 rows × 7 columns = 42 day labels)
    # =========================================================================

    for row in range(6):
        for col in range(7):
            x = 10 + (col * 40)
            y = 90 + (row * 25)

            # Create day label
            lbl = add_label(d, f"lblDay_{row}_{col}", "", x, y, 35, 20)

            # Style the day label
            lbl.Font.Size = 10
            lbl.TextAlign = 2  # fmTextAlignCenter
            lbl.BackStyle = 1  # fmBackStyleOpaque
            lbl.BorderStyle = 1  # fmBorderStyleSingle
            lbl.BorderColor = 12632256  # Light gray

    # =========================================================================
    # Action buttons are now in the top rows (Select in top, Cancel in second)
    # =========================================================================

    # =========================================================================
    # Inject VBA code
    # =========================================================================

    code_path = get_vba_path("frmCalendarPicker.vba", "forms")
    full_content = read_vba_file(code_path)

    # Extract only VBA code (skip form designer section)
    # Form designer section ends with "End", VBA code starts after that
    lines = full_content.split('\n')
    vba_code_start = 0
    for i, line in enumerate(lines):
        if line.strip() == 'Option Explicit':
            vba_code_start = i
            break

    code_content = '\n'.join(lines[vba_code_start:])
    form.CodeModule.AddFromString(code_content)

    print(f"  [OK] Created calendar picker form: {form.Name}")
