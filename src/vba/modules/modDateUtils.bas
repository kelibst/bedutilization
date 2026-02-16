'==============================================================================
' Module: modDateUtils
' Purpose: Centralized date validation, parsing, and formatting utilities
'          Provides locale-independent date handling for all forms
'
' Created: 2026-02-16
' Notes: Eliminates ~100 lines of duplicated date code across forms
'        Works with 64-bit Excel (no DTPicker dependency)
'==============================================================================

Option Explicit

'==============================================================================
' Function: ParseDate
' Purpose:  Parse date string in dd/mm/yyyy format (locale-independent)
'
' Parameters:
'   dateStr  - String in format dd/mm/yyyy (e.g., "14/02/2026")
'   errorMsg - [OUT] Error message if parsing fails
'
' Returns:
'   Variant  - Date value if valid, Empty if invalid
'
' Example:
'   Dim dt As Variant, errMsg As String
'   dt = ParseDate("14/02/2026", errMsg)
'   If IsEmpty(dt) Then MsgBox errMsg
'==============================================================================
Public Function ParseDate(dateStr As String, Optional ByRef errorMsg As String) As Variant
    On Error GoTo ParseError

    ' Clear previous errors
    errorMsg = ""

    ' Validate input
    Dim tempStr As String
    tempStr = Trim(dateStr)
    If tempStr = "" Then
        errorMsg = "Date field is empty."
        ParseDate = Empty
        Exit Function
    End If

    ' Try dd/mm/yyyy format first (locale-independent)
    Dim parts() As String
    parts = Split(tempStr, "/")

    If UBound(parts) = 2 Then
        ' Validate all parts are numeric
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or _
           Not IsNumeric(parts(2)) Then
            errorMsg = "Invalid date format. Use dd/mm/yyyy"
            ParseDate = Empty
            Exit Function
        End If

        Dim d As Long, m As Long, y As Long
        d = CLng(parts(0))
        m = CLng(parts(1))
        y = CLng(parts(2))

        ' Validate year range
        If y < 2020 Or y > 2030 Then
            errorMsg = "Year must be between 2020 and 2030."
            ParseDate = Empty
            Exit Function
        End If

        ' Validate month range
        If m < 1 Or m > 12 Then
            errorMsg = "Month must be between 1 and 12."
            ParseDate = Empty
            Exit Function
        End If

        ' Check valid day for month/year
        Dim maxDay As Long
        maxDay = Day(DateSerial(y, m + 1, 0))
        If d < 1 Or d > maxDay Then
            errorMsg = "Invalid day for " & MonthName(m) & " " & y & _
                       " (max: " & maxDay & ")"
            ParseDate = Empty
            Exit Function
        End If

        ' Create date using explicit DateSerial (locale-independent)
        ParseDate = DateSerial(y, m, d)
        Exit Function
    End If

    ' Fallback to system date parser (locale-dependent)
    If IsDate(dateStr) Then
        ParseDate = CDate(dateStr)
        Exit Function
    End If

    errorMsg = "Invalid date format. Use dd/mm/yyyy"
    ParseDate = Empty
    Exit Function

ParseError:
    errorMsg = "Date parsing error: " & Err.Description
    ParseDate = Empty
End Function

'==============================================================================
' Function: ValidateDate
' Purpose:  Validate that a date value is within acceptable range
'
' Parameters:
'   dt       - Date value to validate
'   errorMsg - [OUT] Error message if validation fails
'
' Returns:
'   Boolean  - True if valid, False if invalid
'
' Example:
'   If Not ValidateDate(admDate, errMsg) Then
'       MsgBox errMsg
'       Exit Sub
'   End If
'==============================================================================
Public Function ValidateDate(dt As Variant, Optional ByRef errorMsg As String) As Boolean
    On Error GoTo ValidateError

    errorMsg = ""

    ' Check if it's a valid date
    If Not IsDate(dt) Then
        errorMsg = "Invalid date value."
        ValidateDate = False
        Exit Function
    End If

    Dim dateValue As Date
    dateValue = CDate(dt)

    ' Check date range (2020-2030)
    If dateValue < DateSerial(2020, 1, 1) Then
        errorMsg = "Date cannot be before January 1, 2020."
        ValidateDate = False
        Exit Function
    End If

    If dateValue > DateSerial(2030, 12, 31) Then
        errorMsg = "Date cannot be after December 31, 2030."
        ValidateDate = False
        Exit Function
    End If

    ValidateDate = True
    Exit Function

ValidateError:
    errorMsg = "Date validation error: " & Err.Description
    ValidateDate = False
End Function

'==============================================================================
' Function: FormatDateDisplay
' Purpose:  Format date for display in user interface (dd/mm/yyyy)
'
' Parameters:
'   dt - Date value to format
'
' Returns:
'   String - Formatted date string or empty string if invalid
'
' Example:
'   txtDate.Value = FormatDateDisplay(Date)  ' "16/02/2026"
'==============================================================================
Public Function FormatDateDisplay(dt As Variant) As String
    On Error Resume Next
    If IsDate(dt) Then
        FormatDateDisplay = Format(CDate(dt), "dd/mm/yyyy")
    Else
        FormatDateDisplay = ""
    End If
    On Error GoTo 0
End Function

'==============================================================================
' Function: FormatDateStorage
' Purpose:  Format date for storage in Excel cells (yyyy-mm-dd)
'
' Parameters:
'   dt - Date value to format
'
' Returns:
'   String - Formatted date string or empty string if invalid
'
' Example:
'   .Cells(1, COL_DATE).Value = FormatDateStorage(admDate)  ' "2026-02-16"
'==============================================================================
Public Function FormatDateStorage(dt As Variant) As String
    On Error Resume Next
    If IsDate(dt) Then
        FormatDateStorage = Format(CDate(dt), "yyyy-mm-dd")
    Else
        FormatDateStorage = ""
    End If
    On Error GoTo 0
End Function

'==============================================================================
' Function: ShowDatePicker
' Purpose:  Display calendar picker and update target TextBox with selected date
'
' Parameters:
'   targetTextBox - Reference to TextBox control to update
'   initialDate   - [Optional] Initial date to show in calendar
'
' Returns:
'   Boolean - True if date was selected, False if cancelled
'
' Example:
'   If ShowDatePicker(txtDate) Then
'       ' Date was selected and txtDate updated
'   End If
'==============================================================================
Public Function ShowDatePicker(targetTextBox As Object, Optional initialDate As Variant) As Boolean
    On Error GoTo ShowPickerError

    ' Default return value
    ShowDatePicker = False

    ' Determine initial date for calendar
    Dim startDate As Date
    If Not IsMissing(initialDate) And IsDate(initialDate) Then
        startDate = CDate(initialDate)
    ElseIf targetTextBox.Value <> "" And IsDate(targetTextBox.Value) Then
        ' Try to parse current textbox value
        Dim errMsg As String
        Dim parsedDate As Variant
        parsedDate = ParseDate(targetTextBox.Value, errMsg)
        If Not IsEmpty(parsedDate) Then
            startDate = CDate(parsedDate)
        Else
            startDate = Date  ' Use today if parsing fails
        End If
    Else
        startDate = Date  ' Use today as default
    End If

    ' Show calendar picker form
    Dim selectedDate As Variant
    On Error Resume Next
    selectedDate = frmCalendarPicker.ShowCalendar(startDate)

    ' Debug: Check if there was an error
    If Err.Number <> 0 Then
        MsgBox "Error calling calendar: " & Err.Description & " (" & Err.Number & ")", vbCritical
        ShowDatePicker = False
        Exit Function
    End If
    On Error GoTo ShowPickerError

    ' Check if a date was selected (not cancelled)
    If Not IsEmpty(selectedDate) Then
        ' Update textbox with selected date
        targetTextBox.Value = FormatDateDisplay(selectedDate)
        ShowDatePicker = True
    Else
        ' User cancelled - no update needed
        ShowDatePicker = False
    End If

    Exit Function

ShowPickerError:
    MsgBox "Error showing calendar picker: " & Err.Description, vbExclamation, "Date Picker Error"
    ShowDatePicker = False
End Function

'==============================================================================
' Function: IsValidDateInput
' Purpose:  Check if TextBox contains valid date and show error if not
'
' Parameters:
'   txt      - TextBox control to validate
'   errorMsg - [OUT] Error message if invalid
'
' Returns:
'   Boolean - True if valid, False if invalid
'
' Example:
'   If Not IsValidDateInput(txtDate, errMsg) Then
'       MsgBox errMsg
'       txtDate.SetFocus
'       Exit Sub
'   End If
'==============================================================================
Public Function IsValidDateInput(txt As Object, Optional ByRef errorMsg As String) As Boolean
    Dim parsedDate As Variant

    ' Parse the date
    parsedDate = ParseDate(txt.Value, errorMsg)

    If IsEmpty(parsedDate) Then
        IsValidDateInput = False
        Exit Function
    End If

    ' Validate the date range
    If Not ValidateDate(parsedDate, errorMsg) Then
        IsValidDateInput = False
        Exit Function
    End If

    IsValidDateInput = True
End Function

'==============================================================================
' Function: GetDateFromString
' Purpose:  Safely parse and validate date string, returning Date or Null
'
' Parameters:
'   dateStr - String to parse
'
' Returns:
'   Variant - Date value if valid, Empty if invalid
'
' Example:
'   Dim dt As Variant
'   dt = GetDateFromString(txtDate.Value)
'   If Not IsEmpty(dt) Then
'       ' Use the date
'   End If
'==============================================================================
Public Function GetDateFromString(dateStr As String) As Variant
    Dim errMsg As String
    Dim parsedDate As Variant

    parsedDate = ParseDate(dateStr, errMsg)

    If IsEmpty(parsedDate) Then
        GetDateFromString = Empty
        Exit Function
    End If

    If Not ValidateDate(parsedDate, errMsg) Then
        GetDateFromString = Empty
        Exit Function
    End If

    GetDateFromString = parsedDate
End Function
