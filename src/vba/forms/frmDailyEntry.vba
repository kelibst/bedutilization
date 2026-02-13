Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private isDirty As Boolean  ' Track if data has been modified
Private isLoading As Boolean  ' Prevent events during load

Private Sub UserForm_Initialize()
    isLoading = True
    isDirty = False

    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Populate month combo
    cmbMonth.AddItem "JANUARY"
    cmbMonth.AddItem "FEBRUARY"
    cmbMonth.AddItem "MARCH"
    cmbMonth.AddItem "APRIL"
    cmbMonth.AddItem "MAY"
    cmbMonth.AddItem "JUNE"
    cmbMonth.AddItem "JULY"
    cmbMonth.AddItem "AUGUST"
    cmbMonth.AddItem "SEPTEMBER"
    cmbMonth.AddItem "OCTOBER"
    cmbMonth.AddItem "NOVEMBER"
    cmbMonth.AddItem "DECEMBER"
    cmbMonth.ListIndex = Month(Date) - 1

    ' Set day spinner (1-31)
    spnDay.Min = 1
    spnDay.Max = 31
    spnDay.Value = Day(Date)
    txtDay.Value = CStr(Day(Date))

    ' Populate ward combo
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i

    ' Select first ward
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Initialize numeric fields to 0
    txtAdmissions.Value = "0"
    txtDischarges.Value = "0"
    txtDeaths.Value = "0"
    txtDeaths24.Value = "0"
    txtTransIn.Value = "0"
    txtTransOut.Value = "0"

    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
    UpdateRecentList
End Sub

Private Sub UpdateRecentList()
    lstRecent.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim i As Long
    For i = startRow To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
           tbl.ListRows(i).Range(1, 1).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 2).Value & " | " & _
                "Adm:" & tbl.ListRows(i).Range(1, 4).Value & _
                " Dis:" & tbl.ListRows(i).Range(1, 5).Value & _
                " Rem:" & tbl.ListRows(i).Range(1, 11).Value
        End If
    Next i
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    ' Calculate actual row (last 10 entries)
    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim actualRow As Long
    actualRow = startRow + lstRecent.ListIndex

    If actualRow > tbl.ListRows.Count Then Exit Sub

    isLoading = True ' Prevent change events while loading
    On Error GoTo DateError

    ' Load the selected entry
    Dim entryDate As Date
    entryDate = CDate(tbl.ListRows(actualRow).Range(1, 1).Value)

    cmbMonth.ListIndex = Month(entryDate) - 1
    spnDay.Value = Day(entryDate)
    txtDay.Value = CStr(Day(entryDate))

    ' Load ward
    Dim wc As String
    wc = tbl.ListRows(actualRow).Range(1, 2).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = wc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load data fields
    txtAdmissions.Value = CStr(tbl.ListRows(actualRow).Range(1, 4).Value)
    txtDischarges.Value = CStr(tbl.ListRows(actualRow).Range(1, 5).Value)
    txtDeaths.Value = CStr(tbl.ListRows(actualRow).Range(1, 6).Value)
    txtDeaths24.Value = CStr(tbl.ListRows(actualRow).Range(1, 7).Value)
    txtTransIn.Value = CStr(tbl.ListRows(actualRow).Range(1, 8).Value)
    txtTransOut.Value = CStr(tbl.ListRows(actualRow).Range(1, 9).Value)

    isLoading = False
    UpdatePrevRemaining
    CalculateRemaining
    lblStatus.Caption = "Loaded entry for editing"
    lblStatus.ForeColor = &H0080FF ' Orange
    isDirty = False
    Exit Sub

DateError:
    isLoading = False
    MsgBox "Error loading entry: Invalid date format. Please contact support.", vbCritical, "Date Error"
    Exit Sub
End Sub

Private Sub cmbWard_Change()
    If isLoading Then Exit Sub

    ' Auto-save if data was modified before changing ward
    If isDirty Then
        If MsgBox("Save changes for current ward before switching?", vbYesNo + vbQuestion) = vbYes Then
            SaveCurrentEntry
        End If
        isDirty = False
    End If

    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub cmbMonth_Change()
    If isLoading Then Exit Sub
    UpdateDateFromControls
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub spnDay_Change()
    If isLoading Then Exit Sub
    txtDay.Value = CStr(spnDay.Value)
    UpdateDateFromControls
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub txtDay_Change()
    If isLoading Then Exit Sub
    On Error Resume Next
    Dim d As Long
    d = CLng(Val(txtDay.Value))
    If d >= 1 And d <= 31 Then
        spnDay.Value = d
    End If
End Sub

Private Sub UpdateDateFromControls()
    ' This just updates the internal date tracking
    ' The actual date is now derived from cmbMonth and spnDay
End Sub

Private Function GetSelectedDate() As Date
    On Error GoTo badDate
    Dim yr As Long, mo As Long, dy As Long
    yr = GetReportYear()
    mo = cmbMonth.ListIndex + 1
    dy = spnDay.Value

    ' Validate day for the month
    Dim maxDay As Long
    maxDay = Day(DateSerial(yr, mo + 1, 0))
    If dy > maxDay Then dy = maxDay

    GetSelectedDate = DateSerial(yr, mo, dy)
    Exit Function
badDate:
    GetSelectedDate = Date
End Function

Private Sub UpdatePrevRemaining()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim prevRem As Long
    prevRem = GetLastRemainingForWard(wc, entryDate)
    lblPrevRemaining.Caption = CStr(prevRem)

    Dim bc As Long
    bc = GetBedComplement(wc)
    lblBedComplement.Caption = CStr(bc)

    CalculateRemaining
End Sub

Private Sub CheckExistingEntry()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim existRow As Long
    existRow = CheckDuplicateDaily(entryDate, wc)

    isLoading = True  ' Prevent dirty flag while loading
    If existRow > 0 Then
        ' Load existing values
        Dim tbl As ListObject
        Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")
        txtAdmissions.Value = CStr(tbl.ListRows(existRow).Range(1, 4).Value)
        txtDischarges.Value = CStr(tbl.ListRows(existRow).Range(1, 5).Value)
        txtDeaths.Value = CStr(tbl.ListRows(existRow).Range(1, 6).Value)
        txtDeaths24.Value = CStr(tbl.ListRows(existRow).Range(1, 7).Value)
        txtTransIn.Value = CStr(tbl.ListRows(existRow).Range(1, 8).Value)
        txtTransOut.Value = CStr(tbl.ListRows(existRow).Range(1, 9).Value)
        lblStatus.Caption = "* Existing entry loaded"
        lblStatus.ForeColor = RGB(200, 100, 0)
    Else
        txtAdmissions.Value = "0"
        txtDischarges.Value = "0"
        txtDeaths.Value = "0"
        txtDeaths24.Value = "0"
        txtTransIn.Value = "0"
        txtTransOut.Value = "0"
            lblStatus.Caption = "New entry"
        lblStatus.ForeColor = RGB(100, 100, 100)
    End If
    isLoading = False
    isDirty = False
    CalculateRemaining
End Sub

Private Sub txtAdmissions_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDischarges_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDeaths_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDeaths24_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTransIn_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTransOut_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub CalculateRemaining()
    On Error Resume Next
    Dim prev As Long, adm As Long, dis As Long
    Dim dth As Long, dth24 As Long, ti As Long, tOut As Long, remVal As Long

    prev = CLng(Val(lblPrevRemaining.Caption))
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    dth24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    ' Formula: Remaining = Prev + Admissions + TransfersIn - (Discharges + Deaths + TransfersOut) - Deaths<24Hrs
    remVal = prev + adm + ti - dis - dth - tOut - dth24
    lblRemaining.Caption = CStr(remVal)

    If remVal < 0 Then
        lblRemaining.ForeColor = RGB(255, 0, 0)
    Else
        lblRemaining.ForeColor = RGB(0, 100, 0)
    End If
End Sub

Private Sub btnPrevDay_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    If spnDay.Value > 1 Then
        spnDay.Value = spnDay.Value - 1
    ElseIf cmbMonth.ListIndex > 0 Then
        cmbMonth.ListIndex = cmbMonth.ListIndex - 1
        spnDay.Value = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
    End If
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnNextDay_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    Dim maxDay As Long
    maxDay = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
    If spnDay.Value < maxDay Then
        spnDay.Value = spnDay.Value + 1
    ElseIf cmbMonth.ListIndex < 11 Then
        cmbMonth.ListIndex = cmbMonth.ListIndex + 1
        spnDay.Value = 1
    End If
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnToday_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    cmbMonth.ListIndex = Month(Date) - 1
    spnDay.Value = Day(Date)
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnSaveNext_Click()
    If SaveCurrentEntry() Then
        isDirty = False
        UpdateRecentList
        ' Move to next ward - preserve current ward index first
        Dim nextIdx As Long
        nextIdx = cmbWard.ListIndex + 1
        If nextIdx < cmbWard.ListCount Then
            isLoading = True  ' Prevent auto-save prompt
            cmbWard.ListIndex = nextIdx
            isLoading = False
            UpdatePrevRemaining
            CheckExistingEntry
        Else
            MsgBox "All wards completed for this date!", vbInformation
            ' Optionally advance to next day
            btnNextDay_Click
        End If
    End If
End Sub

Private Sub btnSaveNextDay_Click()
    If SaveCurrentEntry() Then
        isDirty = False
        UpdateRecentList
        ' Advance date to next day and reset to first ward
        isLoading = True
        Dim maxDay As Long
        maxDay = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
        If spnDay.Value < maxDay Then
            spnDay.Value = spnDay.Value + 1
        ElseIf cmbMonth.ListIndex < 11 Then
            cmbMonth.ListIndex = cmbMonth.ListIndex + 1
            spnDay.Value = 1
        End If
        txtDay.Value = CStr(spnDay.Value)
        ' cmbWard.ListIndex = 0  <-- Removed to persist selection
        isLoading = False
        UpdatePrevRemaining
        CheckExistingEntry
    End If
End Sub

Private Sub btnSaveClose_Click()
    If SaveCurrentEntry() Then
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function SaveCurrentEntry() As Boolean
    SaveCurrentEntry = False

    ' Validate
    If cmbWard.ListIndex < 0 Then
        MsgBox "Please select a ward.", vbExclamation
        Exit Function
    End If

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim adm As Long, dis As Long, dth As Long
    Dim d24 As Long, ti As Long, tOut As Long
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    d24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    SaveDailyEntry entryDate, wc, adm, dis, dth, d24, ti, tOut

    lblStatus.Caption = "Saved: " & wardNames(cmbWard.ListIndex) & " - " & Format(entryDate, "dd/mm/yyyy")
    lblStatus.ForeColor = RGB(0, 128, 0)
    isDirty = False

    SaveCurrentEntry = True
End Function
