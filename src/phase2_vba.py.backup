"""
Phase 2: Inject VBA code into the workbook using win32com
Creates UserForms, standard modules, navigation buttons, and saves as .xlsm
"""
import os
import time
import json
from .config import WorkbookConfig

# ═══════════════════════════════════════════════════════════════════════════════
# VBA FILE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def get_vba_path(filename: str, subfolder: str = "modules") -> str:
    """
    Get absolute path to a VBA source file.
    Args:
        filename: Name of the file (e.g., "modConfig.bas")
        subfolder: Subfolder in src/vba/ (default: "modules")
    """
    # Assuming this script is in src/phase2_vba.py, so we go up one level to src/
    current_dir = os.path.dirname(os.path.abspath(__file__))
    vba_dir = os.path.join(current_dir, "vba", subfolder)
    return os.path.join(vba_dir, filename)

def read_vba_file(path: str) -> str:
    """Read content of a VBA source file."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"VBA source file not found: {path}")
    
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

# ═══════════════════════════════════════════════════════════════════════════════
# USERFORM CREATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def create_daily_entry_form(vbproj):
    """Create the frmDailyEntry UserForm programmatically."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmDailyEntry"
    form.Properties("Caption").Value = "Daily Bed State Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 670

    d = form.Designer

    y = 12  # current Y position

    # Date selection: Month combo + Day spinner + navigation
    _add_label(d, "lblDateLabel", "Date:", 12, y, 40, 18)
    cmb_month = _add_combobox(d, "cmbMonth", 55, y, 100, 20)
    _add_label(d, "lblDayLabel", "Day:", 160, y, 30, 18)
    txt_day = _add_textbox(d, "txtDay", 195, y, 35, 20)
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
    _add_button(d, "btnPrevDay", "< Prev", 260, y, 50, 20)
    _add_button(d, "btnNextDay", "Next >", 315, y, 50, 20)
    _add_button(d, "btnToday", "Today", 370, y, 42, 20)
    y += 28

    # Ward combo
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 120, 18)
    cmb = _add_combobox(d, "cmbWard", 140, y, 240, 22)
    y += 28

    # Previous Remaining
    _add_label(d, "lblPrevRemLabel", "Previous Remaining:", 12, y, 120, 18)
    lbl = _add_label(d, "lblPrevRemaining", "0", 140, y, 80, 18)
    lbl.Font.Bold = True
    lbl.Font.Size = 12
    lbl.ForeColor = 0x006400  # dark green

    _add_label(d, "lblBCLabel", "Bed Complement:", 230, y, 100, 18)
    lbl2 = _add_label(d, "lblBedComplement", "0", 340, y, 60, 18)
    lbl2.Font.Bold = True
    y += 32

    # Separator
    _add_label(d, "lblSep1", "", 12, y, 390, 1).BackColor = 0xC0C0C0
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
        _add_label(d, f"lbl{name}", caption, 12, y, 120, 18)
        _add_textbox(d, name, 140, y, 80, 20)
        y += 28

    # Separator
    _add_label(d, "lblSep2", "", 12, y, 390, 1).BackColor = 0xC0C0C0
    y += 8

    # Calculated Remaining
    _add_label(d, "lblRemLabel", "REMAINING:", 12, y, 120, 20)
    lbl3 = _add_label(d, "lblRemaining", "0", 140, y, 80, 20)
    lbl3.Font.Bold = True
    lbl3.Font.Size = 14
    lbl3.ForeColor = 0x006400
    y += 28

    # Status label
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons row 1
    _add_button(d, "btnSaveNext", "Save && Next Ward", 12, y, 130, 28)
    _add_button(d, "btnSaveNextDay", "Save && Next Day", 150, y, 120, 28)
    _add_button(d, "btnSaveClose", "Save && Close", 278, y, 100, 28)
    y += 32
    # Buttons row 2
    _add_button(d, "btnCancel", "Cancel", 12, y, 90, 28)
    y += 38

    # Recent entries list
    _add_label(d, "lblRecent", "Recent Daily Entries:", 12, y, 150, 18)
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


def create_admission_form(vbproj):
    """Create the frmAdmission UserForm."""
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAdmission"
    form.Properties("Caption").Value = "Patient Admission Record"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 520

    d = form.Designer
    y = 12

    # Date
    _add_label(d, "lblDateLabel", "Admission Date (dd/mm/yyyy):", 12, y, 160, 18)
    _add_textbox(d, "txtDate", 180, y, 120, 20)
    y += 28

    # Ward
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 180, y, 200, 22)
    y += 28

    # Patient ID
    _add_label(d, "lblPIDLabel", "Patient ID / Folder No:", 12, y, 160, 18)
    _add_textbox(d, "txtPatientID", 180, y, 200, 20)
    y += 28

    # Patient Name
    _add_label(d, "lblNameLabel", "Patient Name:", 12, y, 160, 18)
    _add_textbox(d, "txtPatientName", 180, y, 200, 20)
    y += 28

    # Age + Unit
    _add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    _add_textbox(d, "txtAge", 80, y, 60, 20)
    _add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex radio buttons
    _add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS radio buttons
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
    y += 28

    # Status
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    _add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    _add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    _add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Recent admissions list
    _add_label(d, "lblRecent", "Recent Admissions:", 12, y, 150, 18)
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


def create_ages_entry_form(vbproj):
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAgesEntry"
    form.Properties("Caption").Value = "Speed Ages Entry"
    form.Properties("Width").Value = 350
    form.Properties("Height").Value = 520

    d = form.Designer
    y = 12

    # Ward
    _add_label(d, "lblWard", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 100, y, 200, 22)
    y += 32

    # Date
    _add_label(d, "lblDate", "Date:", 12, y, 60, 18)
    _add_textbox(d, "txtDate", 80, y, 120, 20)
    y += 32

    # Divider
    _add_label(d, "lblSep1", "", 12, y, 310, 1).BackColor = 0xC0C0C0
    y += 12

    # Age Entry Area
    _add_label(d, "lblAge", "AGE:", 12, y, 60, 20).Font.Bold = True
    _add_textbox(d, "txtAge", 80, y, 60, 24).Font.Size = 12

    _add_label(d, "lblUnit", "Unit:", 150, y+4, 40, 18)
    _add_combobox(d, "cmbAgeUnit", 195, y, 100, 22)
    y += 38

    # Sex
    _add_label(d, "lblSex", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 20, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 20, "grpSex")
    y += 28

    # Insurance
    _add_label(d, "lblIns", "Health Ins:", 12, y, 65, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 70, 20, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 160, y, 100, 20, "grpNHIS")
    y += 32

    # Status
    lbl = _add_label(d, "lblStatus", "Ready", 12, y, 310, 20)
    lbl.Font.Bold = True
    lbl.ForeColor = 0x808080 # Gray
    y += 24

    # Buttons
    btnSave = _add_button(d, "btnSave", "Save Entry (Enter)", 12, y, 140, 30)
    btnSave.Default = True
    _add_button(d, "btnClose", "Close", 160, y, 100, 30)
    y += 38

    # Recent entries list
    _add_label(d, "lblRecent", "Recent Age Entries:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 310
    lst.Height = 100

    # Inject
    code_path = get_vba_path("frmAgesEntry.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_death_form(vbproj):
    """Create the frmDeath UserForm."""
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmDeath"
    form.Properties("Caption").Value = "Death Record Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 620

    d = form.Designer
    y = 12

    # Date
    _add_label(d, "lblDateLabel", "Date of Death (dd/mm/yyyy):", 12, y, 170, 18)
    _add_textbox(d, "txtDate", 190, y, 120, 20)
    y += 28

    # Ward
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 190, y, 200, 22)
    y += 28

    # Folder Number
    _add_label(d, "lblFolderLabel", "Folder Number:", 12, y, 170, 18)
    _add_textbox(d, "txtFolderNum", 190, y, 200, 20)
    y += 28

    # Name
    _add_label(d, "lblNameLabel", "Name of Deceased:", 12, y, 170, 18)
    _add_textbox(d, "txtName", 190, y, 200, 20)
    y += 28

    # Age + Unit
    _add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    _add_textbox(d, "txtAge", 80, y, 60, 20)
    _add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex
    _add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
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
    _add_label(d, "lblCauseLabel", "Cause of Death:", 12, y, 170, 18)
    cmb = d.Controls.Add("Forms.ComboBox.1")
    cmb.Name = "cmbCause"
    cmb.Left = 190
    cmb.Top = y
    cmb.Width = 200
    cmb.Height = 22
    cmb.Style = 0  # fmStyleDropDownCombo (allows free text)
    y += 32

    # Status
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    _add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    _add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    _add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Recent deaths list
    _add_label(d, "lblRecent", "Recent Deaths:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    code_path = get_vba_path("frmDeath.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_ward_manager_form(vbproj):
    """Create the frmWardManager UserForm."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmWardManager"
    form.Properties("Caption").Value = "Manage Ward Configuration"
    form.Properties("Width").Value = 520
    form.Properties("Height").Value = 400

    d = form.Designer

    y = 12

    # Title
    lbl = _add_label(d, "lblTitle", "Ward Configuration Manager", 12, y, 490, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = _add_label(d, "lblInstructions",
                     "Add or edit wards below. Click 'Export Config' to save to JSON, then rebuild the workbook.",
                     12, y, 490, 30)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 38

    # Ward list (left side)
    _add_label(d, "lblWards", "Wards:", 12, y, 180, 18)
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstWards"
    lst.Left = 12
    lst.Top = y + 22
    lst.Width = 180
    lst.Height = 240

    # Ward details (right side)
    x2 = 205
    _add_label(d, "lblDetails", "Ward Details:", x2, y, 200, 18)
    y2 = y + 22

    # Code
    _add_label(d, "lblCode", "Code:", x2, y2, 80, 18)
    _add_textbox(d, "txtCode", x2 + 105, y2, 100, 20)
    y2 += 26

    # Name
    _add_label(d, "lblName", "Name:", x2, y2, 80, 18)
    _add_textbox(d, "txtName", x2 + 105, y2, 200, 20)
    y2 += 26

    # Bed Complement
    _add_label(d, "lblBeds", "Bed Complement:", x2, y2, 100, 18)
    _add_textbox(d, "txtBeds", x2 + 105, y2, 60, 20)
    y2 += 26

    # Previous Year Remaining
    _add_label(d, "lblPrevRem", "Prev Year Remaining:", x2, y2, 100, 18)
    _add_textbox(d, "txtPrevRemaining", x2 + 105, y2, 60, 20)
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
    _add_label(d, "lblOrder", "Display Order:", x2, y2, 100, 18)
    _add_textbox(d, "txtDisplayOrder", x2 + 105, y2, 60, 20)
    y2 += 32

    # Buttons (right side)
    _add_button(d, "btnNew", "New Ward", x2, y2, 90, 28)
    _add_button(d, "btnSave", "Save", x2 + 95, y2, 90, 28)
    _add_button(d, "btnDelete", "Delete", x2 + 190, y2, 80, 28)

    # Bottom buttons
    y_bottom = 350
    _add_button(d, "btnExport", "Export Config to JSON", 12, y_bottom, 150, 28)
    _add_button(d, "btnClose", "Close", 390, y_bottom, 110, 28)

    code_path = get_vba_path("frmWardManager.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


def create_preferences_manager_form(vbproj):
    """Create the frmPreferencesManager UserForm."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmPreferencesManager"
    form.Properties("Caption").Value = "Hospital Preferences Configuration"
    form.Properties("Width").Value = 500
    form.Properties("Height").Value = 370

    d = form.Designer
    y = 12

    # Title
    lbl = _add_label(d, "lblTitle", "Hospital Preferences", 12, y, 470, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = _add_label(d, "lblInstructions",
                     "Configure hospital-specific preferences. After saving and exporting, rebuild the workbook for changes to take effect.",
                     12, y, 470, 35)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 45

    # Warning frame
    lbl = _add_label(d, "lblWarning",
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
    _add_button(d, "btnSave", "Save to Table", 20, y_buttons, 140, 32)
    _add_button(d, "btnSaveRebuild", "Save & Rebuild", 170, y_buttons, 140, 32)
    _add_button(d, "btnCancel", "Cancel", 360, y_buttons, 120, 32)
    # Bottom row - Export only
    y_buttons += 42
    _add_button(d, "btnExport", "Export to JSON (without rebuild)", 20, y_buttons, 310, 28)

    code_path = get_vba_path("frmPreferencesManager.vba", "forms")
    form.CodeModule.AddFromString(read_vba_file(code_path))


# ─── Helper functions for control creation ───────────────────────────────────

def _add_label(designer, name, caption, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.Label.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_textbox(designer, name, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.TextBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_combobox(designer, name, left, top, width, height, style=0):
    ctrl = designer.Controls.Add("Forms.ComboBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Style = style  # 0=DropDownCombo (default), 2=DropDownList
    return ctrl


def _add_optionbutton(designer, name, caption, left, top, width, height, group_name=None):
    ctrl = designer.Controls.Add("Forms.OptionButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    if group_name:
        ctrl.GroupName = group_name
    return ctrl


def _add_spinner(designer, name, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.SpinButton.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_button(designer, name, caption, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.CommandButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_sheet_button(ws, button_name, cell_range_addr, macro_name):
    """Helper to add a button to a worksheet cell range."""
    cell_range = ws.Range(cell_range_addr)
    left = float(cell_range.Left)
    top = float(cell_range.Top)
    width = float(cell_range.Width)
    height = float(cell_range.Height)

    try:
        # Delete if exists
        ws.Shapes(button_name).Delete()
    except:
        pass

    shp = ws.Shapes.AddShape(5, left, top, width, height)  # 5 = msoShapeRoundedRectangle
    shp.Name = button_name
    
    caption = cell_range.Cells(1, 1).Value
    
    shp.TextFrame.Characters().Text = caption
    shp.TextFrame.Characters().Font.Size = 11
    shp.TextFrame.Characters().Font.Bold = True
    shp.TextFrame.Characters().Font.Color = 16777215  # White
    shp.TextFrame.HorizontalAlignment = -4108  # xlCenter
    shp.TextFrame.VerticalAlignment = -4108
    shp.Fill.ForeColor.RGB = 7884319  # Dark blue
    shp.Line.Visible = False
    shp.OnAction = macro_name



# ═══════════════════════════════════════════════════════════════════════════════
# NAVIGATION BUTTONS ON CONTROL SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def create_nav_buttons(wb):
    """Add navigation shape-buttons to the Control sheet."""
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
    ws.Range("A19").Value = "Manage Wards"  # New button
    ws.Range("A21").Value = "Export Ward Config"  # New button
    ws.Range("A23").Value = "Export Year-End"  # Moved down

    _add_sheet_button(ws, "btnDailyEntry", "Control!A9:C9", "ShowDailyEntry")
    _add_sheet_button(ws, "btnAdmission", "Control!A11:C11", "ShowAdmission")
    _add_sheet_button(ws, "btnDeath", "Control!A13:C13", "ShowDeath")
    _add_sheet_button(ws, "btnAgesEntry", "Control!A15:C15", "ShowAgesEntry")
    _add_sheet_button(ws, "btnRefresh", "Control!A17:C17", "ShowRefreshReports")
    _add_sheet_button(ws, "btnManageWards", "Control!A19:C19", "ShowWardManager")  # New
    _add_sheet_button(ws, "btnExportConfig", "Control!A21:C21", "ExportWardConfig")  # New
    _add_sheet_button(ws, "btnExportYearEnd", "Control!A23:C23", "ExportCarryForward")  # Moved
    _add_sheet_button(ws, "btnPreferences", "Control!A25:C25", "ShowPreferencesInfo")  # New

    # Rebuild button (special orange button)
    _add_sheet_button(ws, "btnRebuild", "Control!A27:C27", "RebuildWorkbookWithPreferences")

    # Diagnostic buttons (row 29, 31, 33 for spacing)
    ws.Range("A29").Value = "Import from Old Workbook"
    ws.Range("A31").Value = "Recalculate All Data"
    ws.Range("A33").Value = "Verify Calculations"
    _add_sheet_button(ws, "btnImport", "Control!A29:C29", "ImportFromOldWorkbook")
    _add_sheet_button(ws, "btnRecalcAll", "Control!A31:C31", "RecalculateAllRows")
    _add_sheet_button(ws, "btnVerify", "Control!A33:C33", "VerifyCalculations")
    # Note: "Fix Date Formats" button removed - date formats now initialized automatically during build


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN INJECTION FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def initialize_date_formats(wb):
    """
    Initialize date column formats for all data tables.
    This ensures date columns are properly formatted from the start.
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


def inject_vba(xlsx_path: str, xlsm_path: str, config: WorkbookConfig):
    """Open xlsx in Excel via COM, inject VBA, save as xlsm."""
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
        excel_processes = [p for p in psutil.process_iter(['name', 'open_files']) if p.info['name'] and 'excel' in p.info['name'].lower()]
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
        excel.Quit()
        time.sleep(1)
