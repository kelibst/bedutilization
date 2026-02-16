"""
UI Helper Functions

Reusable functions for creating UI controls in UserForms and worksheets.
These helpers abstract the win32com API for creating Excel/VBA UI elements.
"""
from typing import Any, Optional


def add_label(designer: Any, name: str, caption: str, left: int, top: int, 
              width: int, height: int) -> Any:
    """
    Add a label control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        caption: Label text
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        
    Returns:
        The created label control
    """
    ctrl = designer.Controls.Add("Forms.Label.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def add_textbox(designer: Any, name: str, left: int, top: int, 
                width: int, height: int) -> Any:
    """
    Add a textbox control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        
    Returns:
        The created textbox control
    """
    ctrl = designer.Controls.Add("Forms.TextBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def add_combobox(designer: Any, name: str, left: int, top: int, 
                 width: int, height: int, style: int = 0) -> Any:
    """
    Add a combobox control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        style: ComboBox style (0=DropDownCombo, 2=DropDownList)
        
    Returns:
        The created combobox control
    """
    ctrl = designer.Controls.Add("Forms.ComboBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Style = style
    return ctrl


def add_optionbutton(designer: Any, name: str, caption: str, left: int, top: int, 
                     width: int, height: int, group_name: Optional[str] = None) -> Any:
    """
    Add an option button (radio button) control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        caption: Button label text
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        group_name: Optional group name for radio button grouping
        
    Returns:
        The created option button control
    """
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


def add_spinner(designer: Any, name: str, left: int, top: int, 
                width: int, height: int) -> Any:
    """
    Add a spinner (spin button) control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        
    Returns:
        The created spinner control
    """
    ctrl = designer.Controls.Add("Forms.SpinButton.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def add_button(designer: Any, name: str, caption: str, left: int, top: int, 
               width: int, height: int) -> Any:
    """
    Add a command button control to a UserForm designer.
    
    Args:
        designer: UserForm designer object
        name: Control name
        caption: Button text
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        
    Returns:
        The created button control
    """
    ctrl = designer.Controls.Add("Forms.CommandButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def add_sheet_button(ws: Any, button_name: str, cell_range_addr: str,
                     macro_name: str) -> None:
    """
    Add a styled button shape to a worksheet cell range.

    Creates a rounded rectangle shape that acts as a button, styled with
    dark blue background and white text, assigned to a macro.

    Args:
        ws: Worksheet object
        button_name: Name for the button shape
        cell_range_addr: Cell range address (e.g., "A9:C9")
        macro_name: Name of the VBA macro to execute on click
    """
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

    # Create rounded rectangle shape
    shp = ws.Shapes.AddShape(5, left, top, width, height)  # 5 = msoShapeRoundedRectangle
    shp.Name = button_name

    # Get caption from cell
    caption = cell_range.Cells(1, 1).Value

    # Style the button
    shp.TextFrame.Characters().Text = caption
    shp.TextFrame.Characters().Font.Size = 11
    shp.TextFrame.Characters().Font.Bold = True
    shp.TextFrame.Characters().Font.Color = 16777215  # White
    shp.TextFrame.HorizontalAlignment = -4108  # xlCenter
    shp.TextFrame.VerticalAlignment = -4108
    shp.Fill.ForeColor.RGB = 7884319  # Dark blue
    shp.Line.Visible = False
    shp.OnAction = macro_name


def add_date_entry_control(designer: Any, name: str, label_text: str, left: int, top: int,
                           label_width: int = 120, textbox_width: int = 100) -> tuple:
    """
    Create complete date entry control set: label + textbox + calendar picker button.

    This is a reusable component that provides:
    - Label describing the date field
    - TextBox for manual date entry (dd/mm/yyyy format)
    - Button to open calendar picker (modDateUtils.ShowDatePicker)

    Args:
        designer: UserForm designer object
        name: Base name for the textbox control (e.g., "txtDate")
        label_text: Label caption (e.g., "Admission Date:")
        left: Left position in points
        top: Top position in points
        label_width: Width of the label in points (default: 120)
        textbox_width: Width of the textbox in points (default: 100)

    Returns:
        tuple: (label_control, textbox_control, button_control)

    Example:
        lbl, txt, btn = add_date_entry_control(d, "txtDate", "Admission Date:", 12, 20)
        # VBA code handler: Private Sub txtDate_picker_Click()
        #                       modDateUtils.ShowDatePicker txtDate
        #                   End Sub
    """
    # Create label
    lbl = add_label(designer, f"{name}_label", label_text, left, top, label_width, 18)

    # Create textbox (positioned after label with 5pt spacing)
    txt_left = left + label_width + 5
    txt = add_textbox(designer, name, txt_left, top, textbox_width, 22)

    # Create calendar picker button (positioned after textbox with 5pt spacing)
    btn_left = txt_left + textbox_width + 5
    btn = add_button(designer, f"{name}_picker", "...", btn_left, top, 24, 22)

    return (lbl, txt, btn)
