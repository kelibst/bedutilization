Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private editingRowIndex As Long  ' 0 = new entry, >0 = editing specific row

Private Sub UserForm_Initialize()
    editingRowIndex = 0  ' Start in new entry mode
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Wards
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Date defaults
    txtDate.Value = modDateUtils.FormatDateDisplay(Date)

    ' Age Units
    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0 ' Default Years

    ' Defaults
    optMale.Value = True
    optInsured.Value = True

    lblStatus.Caption = "Ready"

    ' Initialize date filter controls
    On Error Resume Next
    ' Handle both DTPicker and TextBox
    If TypeName(Me.Controls("dtpFilterDate")) = "DTPicker" Then
        dtpFilterDate.Value = Date
    Else
        dtpFilterDate.Value = Format(Date, "dd/mm/yyyy")
    End If
    On Error GoTo 0

    txtAge.SetFocus
End Sub

Private Sub UpdateRecentList(Optional filterDate As Variant)
    On Error Resume Next
    Application.ScreenUpdating = False
    lstRecent.Clear

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim i As Long
    Dim displayCount As Integer
    displayCount = 0

    ' Check if filtering by date or showing all (last 10)
    If Not IsMissing(filterDate) And Not IsEmpty(filterDate) Then
        ' Filter mode: Show ALL entries matching the selected date
        For i = 1 To tbl.ListRows.Count
            If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) And _
               tbl.ListRows(i).Range(1, 2).Value <> "" Then
                Dim entryDate As Date
                entryDate = CDate(tbl.ListRows(i).Range(1, 2).Value)

                If DateValue(entryDate) = DateValue(filterDate) Then
                    lstRecent.AddItem Format(entryDate, "dd/mm/yyyy") & " | " & _
                        tbl.ListRows(i).Range(1, 6).Value & " | " & _
                        tbl.ListRows(i).Range(1, 7).Value & " " & _
                        tbl.ListRows(i).Range(1, 8).Value & " | " & _
                        tbl.ListRows(i).Range(1, 9).Value & " | " & _
                        tbl.ListRows(i).Range(1, 10).Value
                    displayCount = displayCount + 1
                End If
            End If
        Next i
        lblRecentStatus.Caption = displayCount & " entries on " & Format(filterDate, "dd/mm/yyyy")
    Else
        ' Default mode: Show last 10 entries
        Dim startRow As Long
        startRow = tbl.ListRows.Count - 9
        If startRow < 1 Then startRow = 1

        For i = startRow To tbl.ListRows.Count
            If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) And _
               tbl.ListRows(i).Range(1, 2).Value <> "" Then
                lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 2).Value, "dd/mm/yyyy") & " | " & _
                    tbl.ListRows(i).Range(1, 6).Value & " | " & _
                    tbl.ListRows(i).Range(1, 7).Value & " " & _
                    tbl.ListRows(i).Range(1, 8).Value & " | " & _
                    tbl.ListRows(i).Range(1, 9).Value & " | " & _
                    tbl.ListRows(i).Range(1, 10).Value
                displayCount = displayCount + 1
            End If
        Next i
        lblRecentStatus.Caption = "Last " & displayCount & " entries"
    End If

    Application.ScreenUpdating = True
End Sub

'==============================================================================
' Update Validation Display
' Shows comparison between daily bed-state total and individual admission count
'==============================================================================
Private Sub UpdateValidationDisplay()
    ' Show admission validation status for current date/ward
    On Error Resume Next

    ' Check if lblValidation control exists (it's optional)
    Dim hasLabel As Boolean
    hasLabel = False
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Name = "lblValidation" Then
            hasLabel = True
            Exit For
        End If
    Next ctrl
    If Not hasLabel Then Exit Sub

    ' Need valid date and ward
    If cmbWard.ListIndex < 0 Then Exit Sub
    If Trim(txtDate.Value) = "" Then Exit Sub

    Dim checkDate As Variant
    Dim errMsg As String
    checkDate = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(checkDate) Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim dailyTotal As Long
    Dim individualCount As Long
    Dim validationMsg As String

    If ValidateAdmissionCount(CDate(checkDate), wc, dailyTotal, individualCount, validationMsg) Then
        Me.Controls("lblValidation").Caption = "Daily Total: " & dailyTotal & " | Individual Count: " & individualCount & " [OK]"
        Me.Controls("lblValidation").ForeColor = RGB(0, 128, 0)  ' Green
    Else
        If dailyTotal = 0 And InStr(validationMsg, "No daily bed-state") > 0 Then
            Me.Controls("lblValidation").Caption = "Daily Total: Not entered yet"
            Me.Controls("lblValidation").ForeColor = RGB(128, 128, 128)  ' Gray
        Else
            Me.Controls("lblValidation").Caption = "Daily Total: " & dailyTotal & " | Individual Count: " & individualCount & " [MISMATCH]"
            Me.Controls("lblValidation").ForeColor = RGB(255, 0, 0)  ' Red
        End If
    End If
End Sub

' Event handler for "All Records" option
Private Sub optAllRecords_Click()
    On Error Resume Next
    dtpFilterDate.Enabled = False
    UpdateRecentList
End Sub

' Event handler for "Specific Date" option
Private Sub optSpecificDate_Click()
    On Error Resume Next
    dtpFilterDate.Enabled = True
    UpdateRecentList dtpFilterDate.Value
End Sub

' Event handler for date picker change
Private Sub dtpFilterDate_Change()
    On Error Resume Next
    If optSpecificDate.Value = True Then
        Dim filterVal As Variant
        filterVal = dtpFilterDate.Value
        If IsDate(filterVal) Or (VarType(filterVal) = vbString And Len(filterVal) > 0) Then
            UpdateRecentList CDate(filterVal)
        End If
    End If
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub
    On Error GoTo DateError

    ' Parse the selected item to extract admission date and patient name
    Dim selectedItem As String
    selectedItem = lstRecent.List(lstRecent.ListIndex)

    ' Format: "dd/mm/yyyy | PatientName | Age AgeUnit | Sex | NHIS"
    Dim datePart As String, namePart As String
    Dim firstPipe As Integer
    firstPipe = InStr(selectedItem, "|")

    datePart = Trim(Left(selectedItem, firstPipe - 1))

    ' Find the matching row in the table by date and patient name
    Dim entryDate As Date
    entryDate = CDate(datePart)

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    ' Extract patient name (between first and second |)
    Dim secondPipe As Integer
    secondPipe = InStr(firstPipe + 1, selectedItem, "|")
    namePart = Trim(Mid(selectedItem, firstPipe + 1, secondPipe - firstPipe - 1))

    Dim actualRow As Long
    actualRow = 0
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) Then
            Dim checkDate As Date
            checkDate = CDate(tbl.ListRows(i).Range(1, 2).Value)
            Dim checkName As String
            checkName = tbl.ListRows(i).Range(1, 6).Value

            If DateValue(checkDate) = DateValue(entryDate) And checkName = namePart Then
                actualRow = i
                Exit For
            End If
        End If
    Next i

    If actualRow = 0 Then
        MsgBox "Could not find the selected entry in the table.", vbExclamation
        Exit Sub
    End If

    ' Store the row we're editing
    editingRowIndex = actualRow

    ' Load the selected entry
    Dim dateVal As Variant
    dateVal = tbl.ListRows(actualRow).Range(1, 2).Value

    ' Validate date value
    If IsEmpty(dateVal) Or Not IsDate(dateVal) Then
        MsgBox "Error: Invalid date in selected entry." & vbCrLf & _
               "The date may be corrupted or stored as text." & vbCrLf & _
               "Please rebuild the workbook or contact support.", vbCritical, "Date Error"
        Exit Sub
    End If

    entryDate = CDate(dateVal)

    ' Additional validation - ensure date is not default value
    If entryDate < DateSerial(2020, 1, 1) Or entryDate > DateSerial(2030, 12, 31) Then
        MsgBox "Error: Date out of valid range (2020-2030)." & vbCrLf & _
               "Current value: " & Format(entryDate, "yyyy-mm-dd") & vbCrLf & _
               "Please rebuild the workbook or contact support.", vbCritical, "Date Error"
        Exit Sub
    End If

    txtDate.Value = Format(entryDate, "dd/mm/yyyy")

    ' Load ward
    Dim wc As String
    wc = tbl.ListRows(actualRow).Range(1, COL_ADM_WARD_CODE).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = wc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load age and unit
    txtAge.Value = CStr(tbl.ListRows(actualRow).Range(1, 7).Value)
    cmbAgeUnit.Value = tbl.ListRows(actualRow).Range(1, 8).Value

    ' Load sex
    If tbl.ListRows(actualRow).Range(1, 9).Value = "M" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If

    ' Load NHIS
    If tbl.ListRows(actualRow).Range(1, 10).Value = "Insured" Then
        optInsured.Value = True
    Else
        optNonInsured.Value = True
    End If

    lblStatus.Caption = "Loaded entry for editing"
    lblStatus.ForeColor = RGB(255, 128, 0) ' Orange
    txtAge.SetFocus
    Exit Sub

DateError:
    MsgBox "Error loading entry: Invalid date format. Please contact support.", vbCritical, "Date Error"
    Exit Sub
End Sub

Private Sub btnSave_Click()
    ' Validate
    If cmbWard.ListIndex < 0 Then
        MsgBox "Select Ward", vbExclamation
        Exit Sub
    End If
    If txtAge.Value = "" Or Not IsNumeric(txtAge.Value) Then
        MsgBox "Enter valid Age", vbExclamation
        txtAge.SetFocus
        Exit Sub
    End If

    ' Validate Date
    Dim dt As Variant
    Dim errMsg As String

    ' Parse and validate date using centralized date utils
    dt = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(dt) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Sub
    End If

    If Not modDateUtils.ValidateDate(dt, errMsg) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Sub
    End If

    ' Convert to Date type for use in function call
    dt = CDate(dt)

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim age As Long
    age = CLng(txtAge.Value)
    Dim unit As String
    unit = cmbAgeUnit.Value

    Dim sex As String
    If optMale.Value Then sex = "M" Else sex = "F"

    Dim nhis As String
    If optInsured.Value Then nhis = "Insured" Else nhis = "Non-Insured"

    ' Check if we're editing or creating new
    If editingRowIndex > 0 Then
        ' Edit mode: Update existing row
        UpdateAgesRow editingRowIndex, dt, wc, "-", age, unit, sex, nhis
        editingRowIndex = 0  ' Clear edit mode after save
        lblStatus.Caption = "Updated: " & age & " " & unit & " (" & sex & ", " & nhis & ")"
    Else
        ' New entry mode: Create new row
        ' Use blank patient name for speed entries (no individual patient tracking)
        Application.Run "SaveAdmission", dt, wc, "-", "", age, unit, sex, nhis
        lblStatus.Caption = "Saved: " & age & " " & unit & " (" & sex & ", " & nhis & ")"
    End If

    ' Post-Save Reset
    lblStatus.ForeColor = RGB(0, 128, 0) ' Green

    txtAge.Value = ""
    cmbAgeUnit.ListIndex = 0 ' Reset to Years
    ' Keep persistent selections (Ward, Date, Sex, NHIS)

    UpdateRecentList dt
    UpdateAdmissionTotals CDate(dt)
    txtAge.SetFocus
    Exit Sub

End Sub

Private Sub UpdateAgesRow(rowIndex As Long, admDate As Variant, wardCode As String, _
    patientID As String, age As Long, ageUnit As String, sex As String, nhis As String)
    ' Update existing row instead of creating new one
    On Error GoTo UpdateError

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then
        MsgBox "Error: Invalid row index", vbCritical, "Update Error"
        Exit Sub
    End If

    Dim targetRow As ListRow
    Set targetRow = tbl.ListRows(rowIndex)

    With targetRow.Range
        ' Update all fields (keep existing ID, update other fields)
        If IsDate(admDate) Then
            .Cells(1, COL_ADM_DATE).Value = CDate(admDate)
            .Cells(1, COL_ADM_MONTH).Value = Month(CDate(admDate))
        Else
            .Cells(1, COL_ADM_DATE).Value = admDate
            .Cells(1, COL_ADM_MONTH).Value = 0
        End If
        .Cells(1, COL_ADM_WARD_CODE).Value = wardCode
        .Cells(1, COL_ADM_PATIENT_ID).Value = patientID
        .Cells(1, COL_ADM_PATIENT_NAME).Value = ""            ' Keep blank, consistent with SaveAdmission
        .Cells(1, COL_ADM_AGE).Value = age
        .Cells(1, COL_ADM_AGE_UNIT).Value = ageUnit
        .Cells(1, COL_ADM_SEX).Value = sex
        .Cells(1, COL_ADM_NHIS).Value = nhis
        .Cells(1, COL_ADM_TIMESTAMP).Value = Now
        .Cells(1, COL_ADM_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm"
    End With

    UpdateRecentList CDate(admDate)
    UpdateAdmissionTotals CDate(admDate)
    Exit Sub

UpdateError:
    MsgBox "Error updating age entry: " & Err.Description, vbCritical, "Update Error"
End Sub

Private Sub btnClose_Click()
    editingRowIndex = 0  ' Clear edit mode
    Unload Me
End Sub

'==============================================================================
' Enter key handlers - save from any input control
'==============================================================================
Private Sub txtAge_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        btnSave_Click
    End If
End Sub

Private Sub cmbAgeUnit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        btnSave_Click
    End If
End Sub

Private Sub cmbWard_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        btnSave_Click
    End If
End Sub

Private Sub txtDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        btnSave_Click
    End If
End Sub

'==============================================================================
' Calendar Picker Button Click Handler
'==============================================================================
Private Sub txtDate_picker_Click()
    ' Show calendar picker and update date field
    If modDateUtils.ShowDatePicker(txtDate) Then
        UpdateValidationDisplay
    End If
End Sub

'==============================================================================
' Date Auto-Filter: when entry date changes, filter list and refresh totals
'==============================================================================
Private Sub txtDate_Change()
    On Error Resume Next
    Dim dt As Variant
    Dim errMsg As String
    dt = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If Not IsEmpty(dt) Then
        If modDateUtils.ValidateDate(dt, errMsg) Then
            Dim filterDate As Date
            filterDate = CDate(dt)
            UpdateRecentList filterDate
            UpdateAdmissionTotals filterDate
        End If
    End If
End Sub

'==============================================================================
' Show age entry count vs daily admission record for the given date
'==============================================================================
Private Sub UpdateAdmissionTotals(filterDate As Date)
    On Error Resume Next

    Dim tblAdm As ListObject
    Set tblAdm = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim tblDay As ListObject
    Set tblDay = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    ' Count age entries for this date across all wards
    Dim ageEntries As Long
    ageEntries = 0
    Dim i As Long
    For i = 1 To tblAdm.ListRows.Count
        If Not IsEmpty(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value) Then
            If IsDate(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value) Then
                If DateValue(CDate(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value)) = DateValue(filterDate) Then
                    ageEntries = ageEntries + 1
                End If
            End If
        End If
    Next i

    ' Sum daily admission totals for this date across all wards
    Dim dailyTotal As Long
    dailyTotal = 0
    For i = 1 To tblDay.ListRows.Count
        If Not IsEmpty(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
            If IsDate(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
                If DateValue(CDate(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value)) = DateValue(filterDate) Then
                    dailyTotal = dailyTotal + CLng(Val(tblDay.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value))
                End If
            End If
        End If
    Next i

    lblAdmTotal.Caption = Format(filterDate, "dd/mm/yyyy") & "  —  Entries: " & ageEntries & "  |  Daily record: " & dailyTotal

    If ageEntries > 0 And ageEntries = dailyTotal Then
        lblAdmTotal.ForeColor = RGB(0, 128, 0)    ' Green: match
    ElseIf dailyTotal > 0 And ageEntries <> dailyTotal Then
        lblAdmTotal.ForeColor = RGB(200, 0, 0)    ' Red: mismatch
    Else
        lblAdmTotal.ForeColor = RGB(128, 128, 128) ' Gray: no daily record yet
    End If
End Sub

'==============================================================================
' Validate: compare age entry counts vs daily admission records per ward
'==============================================================================
Private Sub btnValidate_Click()
    Dim dt As Variant
    Dim errMsg As String
    dt = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(dt) Then
        MsgBox "Enter a valid date first.", vbExclamation, "Date Required"
        txtDate.SetFocus
        Exit Sub
    End If

    Dim filterDate As Date
    filterDate = CDate(dt)

    Dim tblAdm As ListObject
    Dim tblDay As ListObject
    Set tblAdm = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")
    Set tblDay = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim wards() As String
    wards = GetWardCodes()

    Dim report As String
    Dim issueCount As Long
    Dim hasAnyData As Boolean
    issueCount = 0
    hasAnyData = False

    report = "Age Entry Validation  —  " & Format(filterDate, "dd/mm/yyyy") & vbCrLf
    report = report & String(50, "-") & vbCrLf

    Dim w As Long
    For w = 0 To UBound(wards)
        Dim wc As String
        wc = wards(w)

        ' Count age entries for this ward + date
        Dim ageCount As Long
        ageCount = 0
        Dim i As Long
        For i = 1 To tblAdm.ListRows.Count
            If Not IsEmpty(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value) Then
                If IsDate(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value) Then
                    If DateValue(CDate(tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value)) = DateValue(filterDate) And _
                       tblAdm.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value = wc Then
                        ageCount = ageCount + 1
                    End If
                End If
            End If
        Next i

        ' Look up daily record for this ward + date
        Dim dailyAdm As Long
        Dim hasDailyRecord As Boolean
        dailyAdm = 0
        hasDailyRecord = False
        For i = 1 To tblDay.ListRows.Count
            If Not IsEmpty(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
                If IsDate(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
                    If DateValue(CDate(tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value)) = DateValue(filterDate) And _
                       tblDay.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value = wc Then
                        dailyAdm = CLng(Val(tblDay.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value))
                        hasDailyRecord = True
                        Exit For
                    End If
                End If
            End If
        Next i

        ' Only include wards with any activity on this date
        If hasDailyRecord Or ageCount > 0 Then
            hasAnyData = True
            If ageCount = dailyAdm Then
                report = report & "  [OK] " & wc & ": " & ageCount & " entries = " & dailyAdm & " admitted" & vbCrLf
            Else
                report = report & "  [!!] " & wc & ": " & ageCount & " entries <> " & dailyAdm & " admitted" & vbCrLf
                issueCount = issueCount + 1
            End If
        End If
    Next w

    report = report & String(50, "-") & vbCrLf

    If Not hasAnyData Then
        MsgBox "No data found for " & Format(filterDate, "dd/mm/yyyy") & "." & vbCrLf & _
               "Ensure both age entries and daily records exist for this date.", _
               vbInformation, "No Data"
        Exit Sub
    End If

    If issueCount = 0 Then
        MsgBox report & "All records match!", vbInformation, "Validation Passed"
    Else
        MsgBox report & issueCount & " ward(s) have inconsistencies. Please verify.", _
               vbExclamation, "Validation Issues Found"
    End If
End Sub
