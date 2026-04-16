Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private isDirty As Boolean  ' Track if data has been modified
Private isLoading As Boolean  ' Prevent events during load
Private isCombinedEmergency As Boolean
Private emergencyWardIndex As Long

Private Function GetPreferenceValue(prefKey As String) As Boolean
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListRows(i).Range(1, 1).Value)) = prefKey Then
            GetPreferenceValue = CBool(tbl.ListRows(i).Range(1, 2).Value)
            Exit Function
        End If
    Next i
    GetPreferenceValue = False
End Function

Private Function IsEmergencySelected() As Boolean
    IsEmergencySelected = (isCombinedEmergency And cmbWard.ListIndex = emergencyWardIndex)
End Function

Private Function GetActualWardCode(comboIndex As Long) As String
    If Not isCombinedEmergency Then
        GetActualWardCode = wardCodes(comboIndex)
        Exit Function
    End If
    Dim origIdx As Long, comboIdx As Long
    comboIdx = 0
    For origIdx = 0 To UBound(wardCodes)
        If wardCodes(origIdx) = "FAE" Then
            ' Skip FAE in combined mode
        Else
            If comboIdx = comboIndex Then
                GetActualWardCode = wardCodes(origIdx)
                Exit Function
            End If
            comboIdx = comboIdx + 1
        End If
    Next origIdx
End Function

Private Sub ToggleEmergencyControls(showCombined As Boolean)
    Dim stdVisible As Boolean
    stdVisible = Not showCombined

    ' Standard controls
    txtAdmissions.Visible = stdVisible
    txtDischarges.Visible = stdVisible
    txtDeaths.Visible = stdVisible
    txtDeaths24.Visible = stdVisible
    txtTransIn.Visible = stdVisible
    txtTransOut.Visible = stdVisible
    lbltxtAdmissions.Visible = stdVisible
    lbltxtDischarges.Visible = stdVisible
    lbltxtDeaths.Visible = stdVisible
    lbltxtDeaths24.Visible = stdVisible
    lbltxtTransIn.Visible = stdVisible
    lbltxtTransOut.Visible = stdVisible
    lblPrevRemaining.Visible = stdVisible
    lblRemaining.Visible = stdVisible
    lblBedComplement.Visible = stdVisible

    ' Emergency combined controls
    lblEmHdrMale.Visible = showCombined
    lblEmHdrFemale.Visible = showCombined
    lblEmHdrTotal.Visible = showCombined

    Dim emSuffixes As Variant
    emSuffixes = Array("Adm", "Dis", "Dth", "D24", "Ti", "To")
    Dim c As Variant
    For Each c In emSuffixes
        Me.Controls("lblEm" & c).Visible = showCombined
        Me.Controls("txt" & c & "M").Visible = showCombined
        Me.Controls("txt" & c & "F").Visible = showCombined
        Me.Controls("lbl" & c & "Total").Visible = showCombined
    Next c

    lblEmPrevRemaining.Visible = showCombined
    lblEmRemaining.Visible = showCombined
    lblEmBedComplement.Visible = showCombined
End Sub

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

    ' Check combined emergency preference
    isCombinedEmergency = GetPreferenceValue("combined_emergency_entry")

    ' Find MAE and FAE in original arrays
    Dim maeFound As Boolean, faeFound As Boolean
    maeFound = False
    faeFound = False
    Dim k As Long
    For k = 0 To UBound(wardCodes)
        If wardCodes(k) = "MAE" Then maeFound = True
        If wardCodes(k) = "FAE" Then faeFound = True
    Next k

    ' Populate ward combo
    emergencyWardIndex = -1
    Dim i As Long
    If isCombinedEmergency And maeFound And faeFound Then
        For i = 0 To UBound(wardNames)
            If wardCodes(i) = "MAE" Then
                cmbWard.AddItem "Emergency"
                emergencyWardIndex = cmbWard.ListCount - 1
            ElseIf wardCodes(i) = "FAE" Then
                ' Skip FAE -- merged into "Emergency"
            Else
                cmbWard.AddItem wardNames(i)
            End If
        Next i
    Else
        isCombinedEmergency = False
        For i = 0 To UBound(wardNames)
            cmbWard.AddItem wardNames(i)
        Next i
    End If

    ' Select first ward
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Initialize standard numeric fields to 0
    txtAdmissions.Value = "0"
    txtDischarges.Value = "0"
    txtDeaths.Value = "0"
    txtDeaths24.Value = "0"
    txtTransIn.Value = "0"
    txtTransOut.Value = "0"

    ' Initialize emergency combined fields to 0
    If isCombinedEmergency Then
        txtAdmM.Value = "0": txtAdmF.Value = "0"
        txtDisM.Value = "0": txtDisF.Value = "0"
        txtDthM.Value = "0": txtDthF.Value = "0"
        txtD24M.Value = "0": txtD24F.Value = "0"
        txtTiM.Value = "0":  txtTiF.Value = "0"
        txtToM.Value = "0":  txtToF.Value = "0"
    End If

    isLoading = False

    ' Initialize date filter controls
    On Error Resume Next
    If TypeName(Me.Controls("dtpFilterDate")) = "DTPicker" Then
        dtpFilterDate.Value = Date
    Else
        dtpFilterDate.Value = Format(Date, "dd/mm/yyyy")
    End If
    On Error GoTo 0

    UpdatePrevRemaining
    CheckExistingEntry
    UpdateRecentList
End Sub

Private Sub UpdateRecentList(Optional filterDate As Variant)
    On Error Resume Next
    Application.ScreenUpdating = False
    lstRecent.Clear

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim i As Long
    Dim displayCount As Integer
    displayCount = 0

    If Not IsMissing(filterDate) And Not IsEmpty(filterDate) Then
        For i = 1 To tbl.ListRows.Count
            If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
               tbl.ListRows(i).Range(1, 1).Value <> "" Then
                Dim entryDate As Date
                entryDate = CDate(tbl.ListRows(i).Range(1, 1).Value)

                If DateValue(entryDate) = DateValue(filterDate) Then
                    lstRecent.AddItem Format(entryDate, "dd/mm/yyyy") & " | " & _
                        tbl.ListRows(i).Range(1, 3).Value & " | " & _
                        "Adm:" & tbl.ListRows(i).Range(1, 4).Value & _
                        " Dis:" & tbl.ListRows(i).Range(1, 5).Value & _
                        " Rem:" & tbl.ListRows(i).Range(1, 11).Value
                    displayCount = displayCount + 1
                End If
            End If
        Next i
        lblRecentStatus.Caption = displayCount & " entries on " & Format(filterDate, "dd/mm/yyyy")
    Else
        Dim startRow As Long
        startRow = tbl.ListRows.Count - 9
        If startRow < 1 Then startRow = 1

        For i = startRow To tbl.ListRows.Count
            If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
               tbl.ListRows(i).Range(1, 1).Value <> "" Then
                lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
                    tbl.ListRows(i).Range(1, 3).Value & " | " & _
                    "Adm:" & tbl.ListRows(i).Range(1, 4).Value & _
                    " Dis:" & tbl.ListRows(i).Range(1, 5).Value & _
                    " Rem:" & tbl.ListRows(i).Range(1, 11).Value
                displayCount = displayCount + 1
            End If
        Next i
        lblRecentStatus.Caption = "Last " & displayCount & " entries"
    End If

    Application.ScreenUpdating = True
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
    isLoading = True

    Dim selectedItem As String
    selectedItem = lstRecent.List(lstRecent.ListIndex)

    Dim datePart As String
    datePart = Left(selectedItem, InStr(selectedItem, "|") - 2)

    Dim firstPipe As Integer, secondPipe As Integer
    firstPipe = InStr(selectedItem, "|")
    secondPipe = InStr(firstPipe + 1, selectedItem, "|")
    Dim wardPart As String
    wardPart = Trim(Mid(selectedItem, firstPipe + 1, secondPipe - firstPipe - 1))

    Dim clickedDate As Date
    clickedDate = CDate(datePart)

    ' Set date controls
    cmbMonth.ListIndex = Month(clickedDate) - 1
    spnDay.Value = Day(clickedDate)
    txtDay.Value = CStr(Day(clickedDate))

    ' Navigate to ward in combo
    If isCombinedEmergency And (wardPart = "MAE" Or wardPart = "FAE") Then
        cmbWard.ListIndex = emergencyWardIndex
    Else
        Dim j As Long
        If isCombinedEmergency Then
            ' Map ward code to combo index (FAE skipped)
            Dim comboIdx As Long
            comboIdx = 0
            For j = 0 To UBound(wardCodes)
                If wardCodes(j) = "FAE" Then
                    ' Skip
                Else
                    If wardCodes(j) = wardPart Then
                        cmbWard.ListIndex = comboIdx
                        Exit For
                    End If
                    comboIdx = comboIdx + 1
                End If
            Next j
        Else
            For j = 0 To UBound(wardCodes)
                If wardCodes(j) = wardPart Then
                    cmbWard.ListIndex = j
                    Exit For
                End If
            Next j
        End If
    End If

    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
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

    ' Toggle emergency combined controls
    If isCombinedEmergency Then
        ToggleEmergencyControls IsEmergencySelected()
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

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    If IsEmergencySelected() Then
        Dim prevM As Long, prevF As Long
        prevM = GetLastRemainingForWard("MAE", entryDate)
        prevF = GetLastRemainingForWard("FAE", entryDate)
        lblEmPrevRemaining.Caption = CStr(prevM + prevF) & " (M: " & prevM & ", F: " & prevF & ")"

        Dim bcM As Long, bcF As Long
        bcM = GetBedComplement("MAE")
        bcF = GetBedComplement("FAE")
        lblEmBedComplement.Caption = CStr(bcM + bcF) & " (M: " & bcM & ", F: " & bcF & ")"
    Else
        Dim wc As String
        wc = GetActualWardCode(cmbWard.ListIndex)

        Dim prevRem As Long
        prevRem = GetLastRemainingForWard(wc, entryDate)
        lblPrevRemaining.Caption = CStr(prevRem)

        Dim bc As Long
        bc = GetBedComplement(wc)
        lblBedComplement.Caption = CStr(bc)
    End If

    CalculateRemaining
End Sub

Private Sub CheckExistingEntry()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    isLoading = True

    If IsEmergencySelected() Then
        Dim maeRow As Long, faeRow As Long
        maeRow = CheckDuplicateDaily(entryDate, "MAE")
        faeRow = CheckDuplicateDaily(entryDate, "FAE")

        Dim tbl As ListObject
        Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

        If maeRow > 0 Then
            txtAdmM.Value = CStr(tbl.ListRows(maeRow).Range(1, 4).Value)
            txtDisM.Value = CStr(tbl.ListRows(maeRow).Range(1, 5).Value)
            txtDthM.Value = CStr(tbl.ListRows(maeRow).Range(1, 6).Value)
            txtD24M.Value = CStr(tbl.ListRows(maeRow).Range(1, 7).Value)
            txtTiM.Value = CStr(tbl.ListRows(maeRow).Range(1, 8).Value)
            txtToM.Value = CStr(tbl.ListRows(maeRow).Range(1, 9).Value)
        Else
            txtAdmM.Value = "0": txtDisM.Value = "0": txtDthM.Value = "0"
            txtD24M.Value = "0": txtTiM.Value = "0": txtToM.Value = "0"
        End If

        If faeRow > 0 Then
            txtAdmF.Value = CStr(tbl.ListRows(faeRow).Range(1, 4).Value)
            txtDisF.Value = CStr(tbl.ListRows(faeRow).Range(1, 5).Value)
            txtDthF.Value = CStr(tbl.ListRows(faeRow).Range(1, 6).Value)
            txtD24F.Value = CStr(tbl.ListRows(faeRow).Range(1, 7).Value)
            txtTiF.Value = CStr(tbl.ListRows(faeRow).Range(1, 8).Value)
            txtToF.Value = CStr(tbl.ListRows(faeRow).Range(1, 9).Value)
        Else
            txtAdmF.Value = "0": txtDisF.Value = "0": txtDthF.Value = "0"
            txtD24F.Value = "0": txtTiF.Value = "0": txtToF.Value = "0"
        End If

        If maeRow > 0 Or faeRow > 0 Then
            lblStatus.Caption = "* Existing emergency entries loaded"
            lblStatus.ForeColor = RGB(200, 100, 0)
        Else
            lblStatus.Caption = "New entry"
            lblStatus.ForeColor = RGB(100, 100, 100)
        End If
    Else
        Dim wc As String
        wc = GetActualWardCode(cmbWard.ListIndex)

        Dim existRow As Long
        existRow = CheckDuplicateDaily(entryDate, wc)

        If existRow > 0 Then
            Dim tbl2 As ListObject
            Set tbl2 = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")
            txtAdmissions.Value = CStr(tbl2.ListRows(existRow).Range(1, 4).Value)
            txtDischarges.Value = CStr(tbl2.ListRows(existRow).Range(1, 5).Value)
            txtDeaths.Value = CStr(tbl2.ListRows(existRow).Range(1, 6).Value)
            txtDeaths24.Value = CStr(tbl2.ListRows(existRow).Range(1, 7).Value)
            txtTransIn.Value = CStr(tbl2.ListRows(existRow).Range(1, 8).Value)
            txtTransOut.Value = CStr(tbl2.ListRows(existRow).Range(1, 9).Value)
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
    End If

    isLoading = False
    isDirty = False
    CalculateRemaining
End Sub

' Standard field change handlers
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

' Emergency combined field change handlers
Private Sub txtAdmM_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtAdmF_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDisM_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDisF_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDthM_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDthF_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtD24M_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtD24F_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTiM_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTiF_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtToM_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtToF_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub CalculateRemaining()
    On Error Resume Next

    If IsEmergencySelected() Then
        Dim entryDate As Date
        entryDate = GetSelectedDate()

        Dim prevM As Long, prevF As Long
        prevM = GetLastRemainingForWard("MAE", entryDate)
        prevF = GetLastRemainingForWard("FAE", entryDate)

        ' Male remaining
        Dim mRem As Long
        mRem = prevM + CLng(Val(txtAdmM.Value)) + CLng(Val(txtTiM.Value)) _
             - CLng(Val(txtDisM.Value)) - CLng(Val(txtDthM.Value)) _
             - CLng(Val(txtToM.Value)) - CLng(Val(txtD24M.Value))

        ' Female remaining
        Dim fRem As Long
        fRem = prevF + CLng(Val(txtAdmF.Value)) + CLng(Val(txtTiF.Value)) _
             - CLng(Val(txtDisF.Value)) - CLng(Val(txtDthF.Value)) _
             - CLng(Val(txtToF.Value)) - CLng(Val(txtD24F.Value))

        Dim totalRem As Long
        totalRem = mRem + fRem

        lblEmRemaining.Caption = CStr(totalRem) & " (M: " & mRem & ", F: " & fRem & ")"
        If totalRem < 0 Then
            lblEmRemaining.ForeColor = RGB(255, 0, 0)
        Else
            lblEmRemaining.ForeColor = RGB(0, 100, 0)
        End If

        ' Update per-field totals
        lblAdmTotal.Caption = CStr(CLng(Val(txtAdmM.Value)) + CLng(Val(txtAdmF.Value)))
        lblDisTotal.Caption = CStr(CLng(Val(txtDisM.Value)) + CLng(Val(txtDisF.Value)))
        lblDthTotal.Caption = CStr(CLng(Val(txtDthM.Value)) + CLng(Val(txtDthF.Value)))
        lblD24Total.Caption = CStr(CLng(Val(txtD24M.Value)) + CLng(Val(txtD24F.Value)))
        lblTiTotal.Caption = CStr(CLng(Val(txtTiM.Value)) + CLng(Val(txtTiF.Value)))
        lblToTotal.Caption = CStr(CLng(Val(txtToM.Value)) + CLng(Val(txtToF.Value)))
    Else
        Dim prev As Long, adm As Long, dis As Long
        Dim dth As Long, dth24 As Long, ti As Long, tOut As Long, remVal As Long

        prev = CLng(Val(lblPrevRemaining.Caption))
        adm = CLng(Val(txtAdmissions.Value))
        dis = CLng(Val(txtDischarges.Value))
        dth = CLng(Val(txtDeaths.Value))
        dth24 = CLng(Val(txtDeaths24.Value))
        ti = CLng(Val(txtTransIn.Value))
        tOut = CLng(Val(txtTransOut.Value))

        remVal = prev + adm + ti - dis - dth - tOut - dth24
        lblRemaining.Caption = CStr(remVal)

        If remVal < 0 Then
            lblRemaining.ForeColor = RGB(255, 0, 0)
        Else
            lblRemaining.ForeColor = RGB(0, 100, 0)
        End If
    End If
End Sub

Private Sub btnPrevDay_Click()
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
        Dim nextIdx As Long
        nextIdx = cmbWard.ListIndex + 1
        If nextIdx < cmbWard.ListCount Then
            isLoading = True
            cmbWard.ListIndex = nextIdx
            isLoading = False
            If isCombinedEmergency Then
                ToggleEmergencyControls IsEmergencySelected()
            End If
            UpdatePrevRemaining
            CheckExistingEntry
        Else
            MsgBox "All wards completed for this date!", vbInformation
            btnNextDay_Click
        End If
    End If
End Sub

Private Sub btnSaveNextDay_Click()
    If SaveCurrentEntry() Then
        isDirty = False
        UpdateRecentList
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

    If cmbWard.ListIndex < 0 Then
        MsgBox "Please select a ward.", vbExclamation
        Exit Function
    End If

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    If IsEmergencySelected() Then
        ' Save TWO rows: one for MAE, one for FAE
        Dim admM As Long, disM As Long, dthM As Long, d24M As Long, tiM As Long, toM As Long
        admM = CLng(Val(txtAdmM.Value))
        disM = CLng(Val(txtDisM.Value))
        dthM = CLng(Val(txtDthM.Value))
        d24M = CLng(Val(txtD24M.Value))
        tiM = CLng(Val(txtTiM.Value))
        toM = CLng(Val(txtToM.Value))

        Dim admF As Long, disF As Long, dthF As Long, d24F As Long, tiF As Long, toF As Long
        admF = CLng(Val(txtAdmF.Value))
        disF = CLng(Val(txtDisF.Value))
        dthF = CLng(Val(txtDthF.Value))
        d24F = CLng(Val(txtD24F.Value))
        tiF = CLng(Val(txtTiF.Value))
        toF = CLng(Val(txtToF.Value))

        SaveDailyEntry entryDate, "MAE", admM, disM, dthM, d24M, tiM, toM
        SaveDailyEntry entryDate, "FAE", admF, disF, dthF, d24F, tiF, toF

        lblStatus.Caption = "Saved: Emergency (M+F) - " & Format(entryDate, "dd/mm/yyyy")
        lblStatus.ForeColor = RGB(0, 128, 0)
    Else
        Dim adm As Long, dis As Long, dth As Long
        Dim d24 As Long, ti As Long, tOut As Long
        adm = CLng(Val(txtAdmissions.Value))
        dis = CLng(Val(txtDischarges.Value))
        dth = CLng(Val(txtDeaths.Value))
        d24 = CLng(Val(txtDeaths24.Value))
        ti = CLng(Val(txtTransIn.Value))
        tOut = CLng(Val(txtTransOut.Value))

        Dim wc As String
        wc = GetActualWardCode(cmbWard.ListIndex)

        SaveDailyEntry entryDate, wc, adm, dis, dth, d24, ti, tOut

        lblStatus.Caption = "Saved: " & cmbWard.List(cmbWard.ListIndex) & " - " & Format(entryDate, "dd/mm/yyyy")
        lblStatus.ForeColor = RGB(0, 128, 0)
    End If

    isDirty = False
    SaveCurrentEntry = True
End Function
