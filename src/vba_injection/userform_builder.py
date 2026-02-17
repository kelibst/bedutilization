"""
UserForm Builder

Functions for creating and configuring UserForms programmatically.
Each function creates a complete UserForm with all controls and injects the VBA code.
"""
from typing import Any
from .ui_helpers import (
    add_label, add_textbox, add_combobox, add_optionbutton,
    add_button, add_spinner, add_date_entry_control, add_listbox
)
from .utils import get_vba_path, read_vba_file
from .calendar_form_builder import create_calendar_picker_form


def add_date_filter_controls(d: Any, y: int, width: int = 390) -> int:
    """
    Add date filter controls for recent list filtering.

    Adds:
    - Label "Filter by Date:"
    - Option button "All Records" (default)
    - Option button "Specific Date"
    - DTPicker control for date selection
    - Status label showing entry count

    Args:
        d: Form designer object
        y: Current Y position
        width: Width of the filter area

    Returns:
        New Y position after adding controls
    """
    # Divider line
    sep = add_label(d, "lblFilterSep", "", 12, y, width, 1)
    sep.BackColor = 0xC0C0C0
    y += 8

    # Filter label
    add_label(d, "lblFilterLabel", "Filter by Date:", 12, y, 100, 18)
    y += 20

    # Option buttons for filter mode
    opt1 = add_optionbutton(d, "optAllRecords", "All Records", 12, y, 120, 18)
    opt1.Value = True  # Default selected
    opt2 = add_optionbutton(d, "optSpecificDate", "Specific Date:", 140, y, 110, 18)
    y += 24

    # Date control - try DTPicker first, fallback to TextBox
    # Always use "dtpFilterDate" as name for consistency with VBA code
    try:
        dtp = d.Controls.Add("MSComCtl2.DTPicker.2")
        dtp.Name = "dtpFilterDate"
        dtp.Left = 140
        dtp.Top = y
        dtp.Width = 120
        dtp.Height = 22
        # Format: dd/mm/yyyy
        dtp.Format = 3  # dtpCustom
        dtp.CustomFormat = "dd/MM/yyyy"
        dtp.Enabled = False  # Disabled by default (All Records is selected)
    except Exception:
        # Fallback to TextBox if DTPicker not available
        # Use same name "dtpFilterDate" for VBA compatibility
        txt = add_textbox(d, "dtpFilterDate", 140, y, 120, 22)
        add_label(d, "lblDateFormat", "(dd/mm/yyyy)", 265, y+2, 80, 18)

    y += 28

    # Status label (shows count of entries)
    lbl = add_label(d, "lblRecentStatus", "Last 10 entries", 12, y, width, 18)
    lbl.ForeColor = 0x808080  # Gray text
    y += 20

    return y


def create_daily_entry_form(vbproj: Any) -> None:
    """
    Create the frmDailyEntry UserForm programmatically.
    
    This form allows users to enter daily bed state data for each ward,
    including admissions, discharges, deaths, and transfers.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmDailyEntry"
    form.Properties("Caption").Value = "Daily Bed State Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 670

    d = form.Designer

    y = 12  # current Y position

    # Date selection: Month combo + Day spinner + navigation
    add_label(d, "lblDateLabel", "Date:", 12, y, 40, 18)
    cmb_month = add_combobox(d, "cmbMonth", 55, y, 100, 20)
    add_label(d, "lblDayLabel", "Day:", 160, y, 30, 18)
    txt_day = add_textbox(d, "txtDay", 195, y, 35, 20)
    # SpinButton for day
    spn = d.Controls.Add("Forms.SpinButton.1")
    spn.Name = "spnDay"
    spn.Left = 232
    spn.Top = y
    spn.Width = 18
    spn.Height = 20
    spn.Min = 1
    spn.Max = 31
    # Navigation buttons
    add_button(d, "btnPrevDay", "< Prev", 260, y, 50, 20)
    add_button(d, "btnNextDay", "Next >", 315, y, 50, 20)
    add_button(d, "btnToday", "Today", 370, y, 42, 20)
    y += 28

    # Ward combo
    add_label(d, "lblWardLabel", "Ward:", 12, y, 120, 18)
    cmb = add_combobox(d, "cmbWard", 140, y, 240, 22)
    y += 28

    # Previous Remaining
    add_label(d, "lblPrevRemLabel", "Previous Remaining:", 12, y, 120, 18)
    lbl = add_label(d, "lblPrevRemaining", "0", 140, y, 80, 18)
    lbl.Font.Bold = True
    lbl.Font.Size = 12
    lbl.ForeColor = 0x006400  # dark green

    add_label(d, "lblBCLabel", "Bed Complement:", 230, y, 100, 18)
    lbl2 = add_label(d, "lblBedComplement", "0", 340, y, 60, 18)
    lbl2.Font.Bold = True
    y += 32

    # Separator
    add_label(d, "lblSep1", "", 12, y, 390, 1).BackColor = 0xC0C0C0
    y += 8

    # Numeric fields
    fields = [
        ("txtAdmissions", "Admissions:"),
        ("txtDischarges", "Discharges:"),
        ("txtDeaths", "Deaths:"),
        ("txtDeaths24", "Deaths < 24Hrs:"),
        ("txtTransIn", "Transfers In:"),
        ("txtTransOut", "Transfers Out:"),
    ]
    for name, caption in fields:
        add_label(d, f"lbl{name}", caption, 12, y, 120, 18)
        add_textbox(d, name, 140, y, 80, 20)
        y += 28

    # Separator
    add_label(d, "lblSep2", "", 12, y, 390, 1).BackColor = 0xC0C0C0
    y += 8

    # Calculated Remaining
    add_label(d, "lblRemLabel", "REMAINING:", 12, y, 120, 20)
    lbl3 = add_label(d, "lblRemaining", "0", 140, y, 80, 20)
    lbl3.Font.Bold = True
    lbl3.Font.Size = 14
    lbl3.ForeColor = 0x006400
    y += 28

    # Status label
    add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons row 1
    add_button(d, "btnSaveNext", "Save && Next Ward", 12, y, 130, 28)
    add_button(d, "btnSaveNextDay", "Save && Next Day", 150, y, 120, 28)
    add_button(d, "btnSaveClose", "Save && Close", 278, y, 100, 28)
    y += 32
    # Buttons row 2
    add_button(d, "btnCancel", "Cancel", 12, y, 90, 28)
    y += 38

    # Date filter controls
    y = add_date_filter_controls(d, y, 390)

    # Recent entries list
    add_label(d, "lblRecent", "Recent Daily Entries:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    # Inject code
    code_path = get_vba_path("frmDailyEntry.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_admission_form(vbproj: Any) -> None:
    """
    Create the frmAdmission UserForm.
    
    This form allows users to record individual patient admission details.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAdmission"
    form.Properties("Caption").Value = "Patient Admission Record"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 520

    d = form.Designer
    y = 12

    # Date with calendar picker
    lbl, txt, btn = add_date_entry_control(d, "txtDate", "Admission Date:", 12, y, label_width=120, textbox_width=100)
    y += 28

    # Ward
    add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    add_combobox(d, "cmbWard", 180, y, 200, 22)
    y += 28

    # Patient ID
    add_label(d, "lblPIDLabel", "Patient ID / Folder No:", 12, y, 160, 18)
    add_textbox(d, "txtPatientID", 180, y, 200, 20)
    y += 28

    # Patient Name
    add_label(d, "lblNameLabel", "Patient Name:", 12, y, 160, 18)
    add_textbox(d, "txtPatientName", 180, y, 200, 20)
    y += 28

    # Age + Unit
    add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    add_textbox(d, "txtAge", 80, y, 60, 20)
    add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex radio buttons
    add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS radio buttons
    add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
    y += 28

    # Status
    add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Date filter controls
    y = add_date_filter_controls(d, y, 380)

    # Recent admissions list
    add_label(d, "lblRecent", "Recent Admissions:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 380
    lst.Height = 100
    
    # Inject code
    code_path = get_vba_path("frmAdmission.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_ages_entry_form(vbproj: Any) -> None:
    """
    Create the frmAgesEntry UserForm.
    
    This form provides a faster way to enter age group data without
    entering full patient details.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAgesEntry"
    form.Properties("Caption").Value = "Speed Ages Entry"
    form.Properties("Width").Value = 350
    form.Properties("Height").Value = 560

    d = form.Designer
    y = 12

    # Ward
    add_label(d, "lblWard", "Ward:", 12, y, 60, 18)
    add_combobox(d, "cmbWard", 100, y, 200, 22)
    y += 32

    # Date with calendar picker
    lbl, txt, btn = add_date_entry_control(d, "txtDate", "Date:", 12, y, label_width=60, textbox_width=100)
    y += 32

    # Divider
    add_label(d, "lblSep1", "", 12, y, 310, 1).BackColor = 0xC0C0C0
    y += 12

    # Age Entry Area
    add_label(d, "lblAge", "AGE:", 12, y, 60, 20).Font.Bold = True
    add_textbox(d, "txtAge", 80, y, 60, 24).Font.Size = 12

    add_label(d, "lblUnit", "Unit:", 150, y+4, 40, 18)
    add_combobox(d, "cmbAgeUnit", 195, y, 100, 22)
    y += 38

    # Sex
    add_label(d, "lblSex", "Sex:", 12, y, 60, 18)
    add_optionbutton(d, "optMale", "Male", 80, y, 60, 20, "grpSex")
    add_optionbutton(d, "optFemale", "Female", 150, y, 70, 20, "grpSex")
    y += 28

    # Insurance
    add_label(d, "lblIns", "Health Ins:", 12, y, 65, 18)
    add_optionbutton(d, "optInsured", "Insured", 80, y, 70, 20, "grpNHIS")
    add_optionbutton(d, "optNonInsured", "Non-Insured", 160, y, 100, 20, "grpNHIS")
    y += 32

    # Status
    lbl = add_label(d, "lblStatus", "Ready", 12, y, 310, 20)
    lbl.Font.Bold = True
    lbl.ForeColor = 0x808080 # Gray
    y += 24

    # Buttons
    btnSave = add_button(d, "btnSave", "Save Entry (Enter)", 12, y, 140, 30)
    btnSave.Default = True
    add_button(d, "btnClose", "Close", 160, y, 100, 30)
    y += 38

    # Date filter controls
    y = add_date_filter_controls(d, y, 310)

    # Totals summary (auto-updated when entry date changes)
    lbl = add_label(d, "lblAdmTotal", "Enter a date above to see totals", 12, y, 310, 18)
    lbl.ForeColor = 0x808080  # Gray
    y += 24

    # Recent entries header + Validate button on same row
    add_label(d, "lblRecent", "Recent Age Entries:", 12, y, 160, 18)
    add_button(d, "btnValidate", "Validate Month", 198, y, 110, 20)
    y += 22
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 310
    lst.Height = 100

    # Inject
    code_path = get_vba_path("frmAgesEntry.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_death_form(vbproj: Any) -> None:
    """
    Create the frmDeath UserForm.
    
    This form allows users to record individual death details.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmDeath"
    form.Properties("Caption").Value = "Death Record Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 620

    d = form.Designer
    y = 12

    # Date with calendar picker
    lbl, txt, btn = add_date_entry_control(d, "txtDate", "Date of Death:", 12, y, label_width=120, textbox_width=100)
    y += 28

    # Ward
    add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    add_combobox(d, "cmbWard", 190, y, 200, 22)
    y += 28

    # Folder Number
    add_label(d, "lblFolderLabel", "Folder Number:", 12, y, 170, 18)
    add_textbox(d, "txtFolderNum", 190, y, 200, 20)
    y += 28

    # Name
    add_label(d, "lblNameLabel", "Name of Deceased:", 12, y, 170, 18)
    add_textbox(d, "txtName", 190, y, 200, 20)
    y += 28

    # Age + Unit
    add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    add_textbox(d, "txtAge", 80, y, 60, 20)
    add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex
    add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS
    add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
    y += 28

    # Death within 24hrs checkbox
    chk = d.Controls.Add("Forms.CheckBox.1")
    chk.Name = "chkWithin24"
    chk.Caption = "Death within 24 hours of admission"
    chk.Left = 12
    chk.Top = y
    chk.Width = 250
    chk.Height = 18
    y += 28

    # Cause of Death
    add_label(d, "lblCauseLabel", "Cause of Death:", 12, y, 170, 18)
    cmb = d.Controls.Add("Forms.ComboBox.1")
    cmb.Name = "cmbCause"
    cmb.Left = 190
    cmb.Top = y
    cmb.Width = 200
    cmb.Height = 22
    cmb.Style = 0  # fmStyleDropDownCombo (allows free text)
    y += 32

    # Status
    add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Date filter controls
    y = add_date_filter_controls(d, y, 390)

    # Recent deaths list
    add_label(d, "lblRecent", "Recent Deaths:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    code_path = get_vba_path("frmDeath.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_ward_manager_form(vbproj: Any) -> None:
    """
    Create the frmWardManager UserForm.
    
    This form allows users to add, edit, and delete ward configurations.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmWardManager"
    form.Properties("Caption").Value = "Manage Ward Configuration"
    form.Properties("Width").Value = 520
    form.Properties("Height").Value = 400

    d = form.Designer

    y = 12

    # Title
    lbl = add_label(d, "lblTitle", "Ward Configuration Manager", 12, y, 490, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = add_label(d, "lblInstructions",
                     "Add or edit wards below. Click 'Export Config' to save to JSON, then rebuild the workbook.",
                     12, y, 490, 30)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 38

    # Ward list (left side)
    add_label(d, "lblWards", "Wards:", 12, y, 180, 18)
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstWards"
    lst.Left = 12
    lst.Top = y + 22
    lst.Width = 180
    lst.Height = 240

    # Ward details (right side)
    x2 = 205
    add_label(d, "lblDetails", "Ward Details:", x2, y, 200, 18)
    y2 = y + 22

    # Code
    add_label(d, "lblCode", "Code:", x2, y2, 80, 18)
    add_textbox(d, "txtCode", x2 + 105, y2, 100, 20)
    y2 += 26

    # Name
    add_label(d, "lblName", "Name:", x2, y2, 80, 18)
    add_textbox(d, "txtName", x2 + 105, y2, 200, 20)
    y2 += 26

    # Bed Complement
    add_label(d, "lblBeds", "Bed Complement:", x2, y2, 100, 18)
    add_textbox(d, "txtBeds", x2 + 105, y2, 60, 20)
    y2 += 26

    # Previous Year Remaining
    add_label(d, "lblPrevRem", "Prev Year Remaining:", x2, y2, 100, 18)
    add_textbox(d, "txtPrevRemaining", x2 + 105, y2, 60, 20)
    y2 += 26

    # Emergency checkbox
    chk = d.Controls.Add("Forms.CheckBox.1")
    chk.Name = "chkEmergency"
    chk.Caption = "Emergency Ward"
    chk.Left = x2
    chk.Top = y2
    chk.Width = 120
    chk.Height = 18
    y2 += 26

    # Display Order
    add_label(d, "lblOrder", "Display Order:", x2, y2, 100, 18)
    add_textbox(d, "txtDisplayOrder", x2 + 105, y2, 60, 20)
    y2 += 32

    # Buttons (right side)
    add_button(d, "btnNew", "New Ward", x2, y2, 90, 28)
    add_button(d, "btnSave", "Save", x2 + 95, y2, 90, 28)
    add_button(d, "btnDelete", "Delete", x2 + 190, y2, 80, 28)

    # Bottom buttons
    y_bottom = 350
    add_button(d, "btnExport", "Export Config to JSON", 12, y_bottom, 150, 28)
    add_button(d, "btnClose", "Close", 390, y_bottom, 110, 28)

    code_path = get_vba_path("frmWardManager.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_preferences_manager_form(vbproj: Any) -> None:
    """
    Create the frmPreferencesManager UserForm.
    
    This form allows users to configure hospital-specific preferences.
    
    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmPreferencesManager"
    form.Properties("Caption").Value = "Hospital Preferences Configuration"
    form.Properties("Width").Value = 500
    form.Properties("Height").Value = 370

    d = form.Designer
    y = 12

    # Title
    lbl = add_label(d, "lblTitle", "Hospital Preferences", 12, y, 470, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = add_label(d, "lblInstructions",
                     "Configure hospital-specific preferences. After saving and exporting, rebuild the workbook for changes to take effect.",
                     12, y, 470, 35)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 45

    # Warning frame
    lbl = add_label(d, "lblWarning",
                     "WARNING: These preferences affect formulas and report structure. Changes require workbook rebuild!",
                     12, y, 470, 30)
    lbl.ForeColor = 0x0000C0  # Dark red
    lbl.WordWrap = True
    lbl.Font.Bold = True
    y += 40

    # Preference checkboxes
    chk1 = d.Controls.Add("Forms.CheckBox.1")
    chk1.Name = "chkShowEmergencyRemaining"
    chk1.Caption = "Show 'Emergency Total Remaining' row in Monthly Summary"
    chk1.Left = 20
    chk1.Top = y
    chk1.Width = 450
    chk1.Height = 18
    y += 30

    chk2 = d.Controls.Add("Forms.CheckBox.1")
    chk2.Name = "chkSubtractDeaths"
    chk2.Caption = "Subtract deaths under 24hrs from monthly admission totals"
    chk2.Left = 20
    chk2.Top = y
    chk2.Width = 450
    chk2.Height = 18
    y += 50

    # Buttons (arranged in 2 rows)
    y_buttons = 240
    # Top row - Primary actions
    add_button(d, "btnSave", "Save to Table", 20, y_buttons, 140, 32)
    add_button(d, "btnSaveRebuild", "Save & Rebuild", 170, y_buttons, 140, 32)
    add_button(d, "btnCancel", "Cancel", 360, y_buttons, 120, 32)
    # Bottom row - Export only
    y_buttons += 42
    add_button(d, "btnExport", "Export to JSON (without rebuild)", 20, y_buttons, 310, 28)

    code_path = get_vba_path("frmPreferencesManager.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_validate_ward_form(vbproj: Any) -> None:
    """
    Create the frmValidateWard UserForm programmatically.

    This form allows users to validate that individual admission entries
    match daily bed-state totals for a specific ward and month.

    Args:
        vbproj: VBProject object from Excel workbook
    """
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmValidateWard"
    form.Properties("Caption").Value = "Validate Ward Admissions"
    form.Properties("Width").Value = 460
    form.Properties("Height").Value = 450

    d = form.Designer
    y = 12  # current Y position

    # Instructions label
    lbl_instr = add_label(d, "lblInstructions",
        "Select a month and ward to validate admission counts:",
        12, y, 420, 18)
    lbl_instr.Font.Bold = True
    y += 28

    # Month selection
    add_label(d, "lblMonth", "Month:", 12, y, 50, 18)
    cmbMonth = add_combobox(d, "cmbMonth", 70, y, 120, 22, style=2)  # DropDownList
    y += 30

    # Ward selection
    add_label(d, "lblWard", "Ward:", 12, y, 50, 18)
    cmbWard = add_combobox(d, "cmbWard", 70, y, 200, 22, style=2)  # DropDownList
    y += 35

    # Action buttons
    add_button(d, "btnValidate", "Validate Month", 12, y, 120, 30)
    add_button(d, "btnExport", "Export Results", 142, y, 120, 30)
    y += 40

    # Results list box
    add_label(d, "lblResults", "Validation Results:", 12, y, 150, 18)
    y += 22

    lstResults = add_listbox(d, "lstResults", 12, y, 420, 220)
    lstResults.ColumnCount = 4
    lstResults.ColumnWidths = "80 pt;60 pt;80 pt;70 pt"
    y += 230

    # Summary label
    lblSummary = add_label(d, "lblSummary", "Select month and ward, then click 'Validate Month'",
                          12, y, 420, 20)
    lblSummary.Font.Bold = True
    lblSummary.TextAlign = 2  # fmTextAlignCenter
    lblSummary.ForeColor = 0x646464  # Gray
    y += 30

    # Close button
    add_button(d, "btnClose", "Close", 320, y, 120, 30)

    # Inject VBA code
    code_path = get_vba_path("frmValidateWard.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))
