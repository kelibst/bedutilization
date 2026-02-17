Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private editingRowIndex As Long  ' 0 = new entry, >0 = editing specific row
Private recentRowIndices() As Long  ' Parallel array storing tblAdmissions row indices for lstRecent

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
    Dim displayCount As Long
    displayCount = 0
    Dim rowDateVal As Variant
    Dim rowWardVal As String

    ' Reset parallel row-index array
    ReDim recentRowIndices(0)

    ' Get current ward filter from combo
    Dim wc As String
    wc = ""
    If cmbWard.ListIndex >= 0 Then wc = wardCodes(cmbWard.ListIndex)

    ' Check if filtering by date or showing all (last 10)
    If Not IsMissing(filterDate) And Not IsEmpty(filterDate) Then
        ' Filter mode: Show ALL entries matching the selected date AND ward
        For i = 1 To tbl.ListRows.Count
            rowDateVal = tbl.ListRows(i).Range(1, COL_ADM_DATE).Value
            If Not IsEmpty(rowDateVal) And rowDateVal <> "" And IsDate(rowDateVal) Then
                rowWardVal = Trim(CStr(tbl.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value))
                If DateValue(CDate(rowDateVal)) = DateValue(filterDate) And _
                   (wc = "" Or rowWardVal = wc) Then
                    lstRecent.AddItem Format(CDate(rowDateVal), "dd/mm/yyyy") & " | " & _
                        rowWardVal & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_AGE).Value & " " & _
                        tbl.ListRows(i).Range(1, COL_ADM_AGE_UNIT).Value & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_SEX).Value & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_NHIS).Value
                    ReDim Preserve recentRowIndices(displayCount)
                    recentRowIndices(displayCount) = i
                    displayCount = displayCount + 1
                End If
            End If
        Next i
        Dim wardLabel As String
        If wc <> "" Then wardLabel = "  " & wc Else wardLabel = ""
        lblRecentStatus.Caption = displayCount & " entries" & wardLabel & "  " & Format(filterDate, "dd/mm/yyyy")
    Else
        ' Default mode: Show last 10 entries for the selected ward
        Dim startRow As Long
        startRow = tbl.ListRows.Count - 9
        If startRow < 1 Then startRow = 1

        For i = startRow To tbl.ListRows.Count
            rowDateVal = tbl.ListRows(i).Range(1, COL_ADM_DATE).Value
            If Not IsEmpty(rowDateVal) And rowDateVal <> "" And IsDate(rowDateVal) Then
                rowWardVal = Trim(CStr(tbl.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value))
                If wc = "" Or rowWardVal = wc Then
                    lstRecent.AddItem Format(CDate(rowDateVal), "dd/mm/yyyy") & " | " & _
                        rowWardVal & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_AGE).Value & " " & _
                        tbl.ListRows(i).Range(1, COL_ADM_AGE_UNIT).Value & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_SEX).Value & " | " & _
                        tbl.ListRows(i).Range(1, COL_ADM_NHIS).Value
                    ReDim Preserve recentRowIndices(displayCount)
                    recentRowIndices(displayCount) = i
                    displayCount = displayCount + 1
                End If
            End If
        Next i
        Dim wardSuffix As String
        If wc <> "" Then wardSuffix = " (" & wc & ")" Else wardSuffix = ""
        lblRecentStatus.Caption = "Last " & displayCount & " entries" & wardSuffix
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
    On Error GoTo LoadError

    ' Use parallel array to get the stored row index directly (no fragile text parsing)
    Dim actualRow As Long
    If lstRecent.ListIndex > UBound(recentRowIndices) Then
        MsgBox "Index out of range. Please refresh the list.", vbExclamation
        Exit Sub
    End If
    actualRow = recentRowIndices(lstRecent.ListIndex)

    If actualRow < 1 Then
        MsgBox "Could not find the selected entry.", vbExclamation
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    If actualRow > tbl.ListRows.Count Then
        MsgBox "Entry no longer exists in the table.", vbExclamation
        Exit Sub
    End If

    ' Store the row we're editing
    editingRowIndex = actualRow

    ' Load date
    Dim dateVal As Variant
    dateVal = tbl.ListRows(actualRow).Range(1, COL_ADM_DATE).Value

    If IsEmpty(dateVal) Or Not IsDate(dateVal) Then
        MsgBox "Error: Invalid date in selected entry." & vbCrLf & _
               "The date may be corrupted. Please rebuild the workbook.", vbCritical, "Date Error"
        Exit Sub
    End If

    Dim entryDate As Date
    entryDate = CDate(dateVal)

    If entryDate < DateSerial(2020, 1, 1) Or entryDate > DateSerial(2030, 12, 31) Then
        MsgBox "Error: Date out of valid range (2020-2030)." & vbCrLf & _
               "Value: " & Format(entryDate, "yyyy-mm-dd"), vbCritical, "Date Error"
        Exit Sub
    End If

    txtDate.Value = Format(entryDate, "dd/mm/yyyy")

    ' Load ward
    Dim rowWc As String
    rowWc = tbl.ListRows(actualRow).Range(1, COL_ADM_WARD_CODE).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = rowWc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load age and unit
    txtAge.Value = CStr(tbl.ListRows(actualRow).Range(1, COL_ADM_AGE).Value)
    cmbAgeUnit.Value = tbl.ListRows(actualRow).Range(1, COL_ADM_AGE_UNIT).Value

    ' Load sex
    If tbl.ListRows(actualRow).Range(1, COL_ADM_SEX).Value = "M" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If

    ' Load NHIS
    If tbl.ListRows(actualRow).Range(1, COL_ADM_NHIS).Value = "Insured" Then
        optInsured.Value = True
    Else
        optNonInsured.Value = True
    End If

    lblStatus.Caption = "Loaded entry for editing"
    lblStatus.ForeColor = RGB(255, 128, 0) ' Orange
    txtAge.SetFocus
    Exit Sub

LoadError:
    MsgBox "Error loading entry: " & Err.Description, vbCritical, "Load Error"
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

'==============================================================================
' Ward Change: refresh list and totals for the newly selected ward
'==============================================================================
Private Sub cmbWard_Change()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub
    Dim dt As Variant
    Dim errMsg As String
    dt = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If Not IsEmpty(dt) Then
        If modDateUtils.ValidateDate(dt, errMsg) Then
            UpdateRecentList CDate(dt)
            UpdateAdmissionTotals CDate(dt)
        End If
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
' Show age entry count vs daily admission record for the selected date AND ward
'==============================================================================
Private Sub UpdateAdmissionTotals(filterDate As Date)
    On Error Resume Next

    If cmbWard.ListIndex < 0 Then
        lblAdmTotal.Caption = "Select a ward to see totals"
        lblAdmTotal.ForeColor = RGB(128, 128, 128)
        Exit Sub
    End If

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim tblAdm As ListObject
    Set tblAdm = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim tblDay As ListObject
    Set tblDay = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim i As Long
    Dim rowDate As Variant

    ' Count age entries for this ward + date
    Dim ageEntries As Long
    ageEntries = 0
    For i = 1 To tblAdm.ListRows.Count
        rowDate = tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value
        If Not IsEmpty(rowDate) And IsDate(rowDate) Then
            If DateValue(CDate(rowDate)) = DateValue(filterDate) And _
               Trim(CStr(tblAdm.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value)) = wc Then
                ageEntries = ageEntries + 1
            End If
        End If
    Next i

    ' Get daily admission total for this ward + date
    Dim dailyTotal As Long
    Dim hasDailyRecord As Boolean
    dailyTotal = 0
    hasDailyRecord = False
    For i = 1 To tblDay.ListRows.Count
        rowDate = tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value
        If Not IsEmpty(rowDate) And IsDate(rowDate) Then
            If DateValue(CDate(rowDate)) = DateValue(filterDate) And _
               Trim(CStr(tblDay.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value)) = wc Then
                dailyTotal = CLng(Val(tblDay.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value))
                hasDailyRecord = True
                Exit For
            End If
        End If
    Next i

    Dim dailyStr As String
    If hasDailyRecord Then dailyStr = CStr(dailyTotal) Else dailyStr = Chr(8212)  ' em dash

    lblAdmTotal.Caption = wc & "  " & Format(filterDate, "dd/mm/yyyy") & _
        "   Entries: " & ageEntries & "  |  Daily: " & dailyStr

    If hasDailyRecord And ageEntries = dailyTotal Then
        lblAdmTotal.ForeColor = RGB(0, 128, 0)    ' Green: match
    ElseIf hasDailyRecord And ageEntries <> dailyTotal Then
        lblAdmTotal.ForeColor = RGB(200, 0, 0)    ' Red: mismatch
    Else
        lblAdmTotal.ForeColor = RGB(128, 128, 128) ' Gray: no daily record yet
    End If
End Sub

'==============================================================================
' Validate Month: 3-check monthly validation for the selected month/ward
'   1. Monthly totals: sum of daily records vs total age entries per ward
'   2. Missing days: days with daily records but zero age entries
'   3. Age anomalies: flag suspicious age values
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
    Dim monthIdx As Long
    monthIdx = Month(filterDate)
    Dim yearIdx As Long
    yearIdx = Year(filterDate)
    Dim monthLabel As String
    monthLabel = Format(filterDate, "mmmm yyyy")

    ' Scope: selected ward or all wards
    Dim selectedWard As String
    selectedWard = ""
    If cmbWard.ListIndex >= 0 Then selectedWard = wardCodes(cmbWard.ListIndex)

    Dim tblAdm As ListObject
    Dim tblDay As ListObject
    Set tblAdm = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")
    Set tblDay = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim wards() As String
    wards = GetWardCodes()

    Dim report As String
    Dim issueCount As Long
    issueCount = 0
    Dim hasAnyData As Boolean
    hasAnyData = False

    Dim wardScope As String
    If selectedWard <> "" Then wardScope = selectedWard Else wardScope = "All Wards"

    report = "Monthly Validation  —  " & monthLabel & "  (" & wardScope & ")" & vbCrLf
    report = report & String(55, "=") & vbCrLf & vbCrLf

    ' ===== CHECK 1: Monthly totals per ward =====
    report = report & "1. MONTHLY TOTALS (Daily records vs Age entries):" & vbCrLf

    Dim w As Long
    Dim i As Long
    Dim vWc As String
    Dim vRowDate As Variant
    Dim totalDailyAll As Long
    Dim totalEntriesAll As Long
    totalDailyAll = 0
    totalEntriesAll = 0

    For w = 0 To UBound(wards)
        vWc = wards(w)
        If selectedWard = "" Or vWc = selectedWard Then
            ' Sum daily admissions for this ward across the month
            Dim monthlyDaily As Long
            monthlyDaily = 0
            Dim wardHasDaily As Boolean
            wardHasDaily = False
            For i = 1 To tblDay.ListRows.Count
                vRowDate = tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value
                If IsDate(vRowDate) Then
                    If Month(CDate(vRowDate)) = monthIdx And Year(CDate(vRowDate)) = yearIdx And _
                       Trim(CStr(tblDay.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value)) = vWc Then
                        monthlyDaily = monthlyDaily + CLng(Val(tblDay.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value))
                        wardHasDaily = True
                    End If
                End If
            Next i

            ' Count age entries for this ward across the month
            Dim monthlyEntries As Long
            monthlyEntries = 0
            For i = 1 To tblAdm.ListRows.Count
                vRowDate = tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value
                If IsDate(vRowDate) Then
                    If Month(CDate(vRowDate)) = monthIdx And Year(CDate(vRowDate)) = yearIdx And _
                       Trim(CStr(tblAdm.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value)) = vWc Then
                        monthlyEntries = monthlyEntries + 1
                    End If
                End If
            Next i

            If wardHasDaily Or monthlyEntries > 0 Then
                hasAnyData = True
                totalDailyAll = totalDailyAll + monthlyDaily
                totalEntriesAll = totalEntriesAll + monthlyEntries
                If monthlyEntries = monthlyDaily Then
                    report = report & "  [OK]  " & vWc & ": Daily=" & monthlyDaily & ", Entries=" & monthlyEntries & vbCrLf
                ElseIf Not wardHasDaily Then
                    report = report & "  [--]  " & vWc & ": No daily records, Entries=" & monthlyEntries & vbCrLf
                    issueCount = issueCount + 1
                Else
                    Dim mDiff As Long
                    mDiff = monthlyEntries - monthlyDaily
                    Dim mDiffStr As String
                    If mDiff > 0 Then mDiffStr = "+" & mDiff Else mDiffStr = CStr(mDiff)
                    report = report & "  [!!]  " & vWc & ": Daily=" & monthlyDaily & ", Entries=" & monthlyEntries & " (" & mDiffStr & ")" & vbCrLf
                    issueCount = issueCount + 1
                End If
            End If
        End If
    Next w

    If Not hasAnyData Then
        MsgBox "No data found for " & monthLabel & "." & vbCrLf & _
               "Ensure both age entries and daily records exist for this month.", _
               vbInformation, "No Data"
        Exit Sub
    End If

    If selectedWard = "" Then
        report = report & "  TOTAL: Daily=" & totalDailyAll & ", Entries=" & totalEntriesAll & vbCrLf
    End If
    report = report & vbCrLf

    ' ===== CHECK 2: Missing days =====
    report = report & "2. MISSING AGE ENTRIES (days with daily records but 0 entries):" & vbCrLf

    Dim missingDays As Long
    missingDays = 0
    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(yearIdx, monthIdx + 1, 0))

    Dim d As Long
    Dim checkDate As Date
    Dim dayDailyTotal As Long
    Dim dayEntryCount As Long
    Dim vDayDate As Variant
    Dim vDayWard As String

    For d = 1 To daysInMonth
        checkDate = DateSerial(yearIdx, monthIdx, d)

        ' Sum daily admissions for this day (filtered by ward)
        dayDailyTotal = 0
        For i = 1 To tblDay.ListRows.Count
            vDayDate = tblDay.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value
            If IsDate(vDayDate) Then
                If DateValue(CDate(vDayDate)) = DateValue(checkDate) Then
                    vDayWard = Trim(CStr(tblDay.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value))
                    If selectedWard = "" Or vDayWard = selectedWard Then
                        dayDailyTotal = dayDailyTotal + CLng(Val(tblDay.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value))
                    End If
                End If
            End If
        Next i

        If dayDailyTotal > 0 Then
            ' Count age entries for this day (filtered by ward)
            dayEntryCount = 0
            For i = 1 To tblAdm.ListRows.Count
                vDayDate = tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value
                If IsDate(vDayDate) Then
                    If DateValue(CDate(vDayDate)) = DateValue(checkDate) Then
                        If selectedWard = "" Or Trim(CStr(tblAdm.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value)) = selectedWard Then
                            dayEntryCount = dayEntryCount + 1
                        End If
                    End If
                End If
            Next i

            If dayEntryCount = 0 Then
                report = report & "  [!!]  " & Format(checkDate, "dd/mm/yyyy") & ": Daily=" & dayDailyTotal & ", Entries=0" & vbCrLf
                missingDays = missingDays + 1
                issueCount = issueCount + 1
            End If
        End If
    Next d

    If missingDays = 0 Then
        report = report & "  [OK]  No missing days" & vbCrLf
    End If
    report = report & vbCrLf

    ' ===== CHECK 3: Age anomalies =====
    report = report & "3. AGE ANOMALIES:" & vbCrLf

    Dim anomalyCount As Long
    anomalyCount = 0
    Dim vAnomDate As Variant
    Dim vAnomWard As String
    Dim anomAgeVal As Long
    Dim anomAgeUnit As String
    Dim anomMsg As String

    For i = 1 To tblAdm.ListRows.Count
        vAnomDate = tblAdm.ListRows(i).Range(1, COL_ADM_DATE).Value
        If IsDate(vAnomDate) Then
            If Month(CDate(vAnomDate)) = monthIdx And Year(CDate(vAnomDate)) = yearIdx Then
                vAnomWard = Trim(CStr(tblAdm.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value))
                If selectedWard = "" Or vAnomWard = selectedWard Then
                    anomAgeVal = CLng(Val(tblAdm.ListRows(i).Range(1, COL_ADM_AGE).Value))
                    anomAgeUnit = CStr(tblAdm.ListRows(i).Range(1, COL_ADM_AGE_UNIT).Value)
                    anomMsg = ""

                    If anomAgeUnit = "Years" Then
                        If anomAgeVal > 110 Then anomMsg = anomAgeVal & " years (unusually old)"
                        If anomAgeVal = 0 Then anomMsg = "0 years (consider Months/Days)"
                    ElseIf anomAgeUnit = "Months" Then
                        If anomAgeVal > 24 Then anomMsg = anomAgeVal & " months (consider Years)"
                        If anomAgeVal = 0 Then anomMsg = "0 months (consider Days)"
                    ElseIf anomAgeUnit = "Days" Then
                        If anomAgeVal > 28 Then anomMsg = anomAgeVal & " days (consider Months)"
                    End If

                    If anomMsg <> "" Then
                        report = report & "  [!]  " & Format(CDate(vAnomDate), "dd/mm/yyyy") & " " & vAnomWard & ": " & anomMsg & vbCrLf
                        anomalyCount = anomalyCount + 1
                    End If
                End If
            End If
        End If
    Next i

    If anomalyCount = 0 Then
        report = report & "  [OK]  No anomalies detected" & vbCrLf
    End If

    report = report & vbCrLf & String(55, "=") & vbCrLf

    If issueCount = 0 Then
        report = report & "All checks passed!"
        MsgBox report, vbInformation, "Validation Passed  —  " & monthLabel
    Else
        report = report & issueCount & " issue(s) require attention."
        MsgBox report, vbExclamation, "Validation Issues  —  " & monthLabel
    End If
End Sub
