Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private editingRowIndex As Long  ' 0 = new entry, >0 = editing specific row

Private Sub UserForm_Initialize()
    editingRowIndex = 0  ' Start in new entry mode
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Load age units FIRST (before ward triggers any change events)
    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0

    ' Now load wards
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    txtDate.Value = modDateUtils.FormatDateDisplay(Date)
    optMale.Value = True
    optInsured.Value = True
    chkWithin24.Value = False

    ' Populate cause of death combo with previous entries
    PopulateCauses

    ' Populate cmbMonth
    cmbMonth.Clear
    For i = 1 To 12
        cmbMonth.AddItem MonthName(i)
    Next i
    cmbMonth.ListIndex = Month(Date) - 1 ' Triggers cmbMonth_Change
End Sub

Private Sub cmbMonth_Change()
    If cmbMonth.ListIndex < 0 Then Exit Sub
    UpdatePendingList
End Sub

Private Sub UpdatePendingList()
    On Error Resume Next
    Application.ScreenUpdating = False
    lstRecent.Clear
    lblMonthStatus.Caption = ""
    lblMonthStatus.ForeColor = RGB(0, 0, 255)

    Dim selectedMonth As Long
    selectedMonth = cmbMonth.ListIndex + 1

    Dim tblDay As ListObject
    Set tblDay = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")
    
    Dim tblDeaths As ListObject
    Set tblDeaths = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    Dim r As Long, rowDate As Date, rowWard As String
    Dim deathsTotal As Long, deathsEntered As Long
    Dim dRow As Long
    Dim displayCount As Integer
    Dim monthHasData As Boolean
    
    displayCount = 0
    monthHasData = False

    For r = 1 To tblDay.ListRows.Count
        If Not IsEmpty(tblDay.ListRows(r).Range(1, 1).Value) Then
            Dim entryMth As Long
            entryMth = CLng(tblDay.ListRows(r).Range(1, 2).Value)
            
            If entryMth = selectedMonth Then
                monthHasData = True
                rowDate = CDate(tblDay.ListRows(r).Range(1, 1).Value)
                rowWard = Trim(CStr(tblDay.ListRows(r).Range(1, 3).Value))
                deathsTotal = CLng(Val(tblDay.ListRows(r).Range(1, 6).Value)) + CLng(Val(tblDay.ListRows(r).Range(1, 7).Value))

                If deathsTotal > 0 Then
                    ' Count how many entered in tblDeaths
                    deathsEntered = 0
                    For dRow = 1 To tblDeaths.ListRows.Count
                        If Not IsEmpty(tblDeaths.ListRows(dRow).Range(1, 2).Value) Then
                            If DateValue(CDate(tblDeaths.ListRows(dRow).Range(1, 2).Value)) = DateValue(rowDate) And _
                               Trim(CStr(tblDeaths.ListRows(dRow).Range(1, 4).Value)) = rowWard Then
                                deathsEntered = deathsEntered + 1
                            End If
                        End If
                    Next dRow

                    If deathsEntered < deathsTotal Then
                        lstRecent.AddItem Format(rowDate, "dd/mm/yyyy") & " | Pending | " & (deathsTotal - deathsEntered) & " missing | " & rowWard
                        displayCount = displayCount + 1
                    End If
                End If
            End If
        End If
    Next r

    If Not monthHasData Then
        lblMonthStatus.Caption = "No daily entries exist for this month yet. Please complete Daily Bed Entry first."
        lblMonthStatus.ForeColor = RGB(255, 0, 0) ' Red
    ElseIf displayCount = 0 Then
        lblMonthStatus.Caption = "All deaths for this month have been fully entered."
        lblMonthStatus.ForeColor = RGB(0, 128, 0) ' Green
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub

    On Error GoTo LoadError

    ' Parse the selected item to extract death date, deceased name, folder number, and ward
    Dim selectedItem As String
    selectedItem = lstRecent.List(lstRecent.ListIndex)

    ' Format: "dd/mm/yyyy | DeceasedName | FolderNumber | Ward"
    Dim datePart As String, namePart As String, folderPart As String, wardPart As String
    Dim firstPipe As Integer, secondPipe As Integer, thirdPipe As Integer
    firstPipe = InStr(selectedItem, "|")
    secondPipe = InStr(firstPipe + 1, selectedItem, "|")
    thirdPipe = InStr(secondPipe + 1, selectedItem, "|")

    datePart = Trim(Left(selectedItem, firstPipe - 1))
    namePart = Trim(Mid(selectedItem, firstPipe + 1, secondPipe - firstPipe - 1))
    folderPart = Trim(Mid(selectedItem, secondPipe + 1, thirdPipe - secondPipe - 1))
    wardPart = Trim(Mid(selectedItem, thirdPipe + 1))
    
    If namePart = "Pending" Then
        ' Redirect for pending entry
        txtDate.Value = datePart
        
        Dim jWard As Long
        For jWard = 0 To UBound(wardCodes)
            If wardCodes(jWard) = wardPart Then
                cmbWard.ListIndex = jWard
                Exit For
            End If
        Next jWard
        
        editingRowIndex = 0
        txtFolderNum.Value = ""
        txtName.Value = ""
        txtAge.Value = ""
        cmbCause.Value = ""
        chkWithin24.Value = False
        lblStatus.Caption = "Ready for new entry (" & wardPart & ")"
        txtFolderNum.SetFocus
        Exit Sub
    End If

    ' Find the matching row in the table by date, deceased name, and folder number
    Dim entryDate As Date
    entryDate = CDate(datePart)

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    Dim actualRow As Long
    actualRow = 0
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) Then
            Dim checkDate As Date
            checkDate = CDate(tbl.ListRows(i).Range(1, 2).Value)
            Dim checkName As String
            checkName = tbl.ListRows(i).Range(1, 6).Value
            Dim checkFolder As String
            checkFolder = tbl.ListRows(i).Range(1, 5).Value

            If DateValue(checkDate) = DateValue(entryDate) And _
               checkName = namePart And checkFolder = folderPart Then
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
    txtDate.Value = Format(tbl.ListRows(actualRow).Range(1, 2).Value, "dd/mm/yyyy")

    ' Load ward
    Dim wc As String
    wc = tbl.ListRows(actualRow).Range(1, 4).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = wc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load patient details
    txtFolderNum.Value = tbl.ListRows(actualRow).Range(1, 5).Value
    txtName.Value = tbl.ListRows(actualRow).Range(1, 6).Value
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

    ' Load within 24 hours flag
    chkWithin24.Value = (tbl.ListRows(actualRow).Range(1, COL_DEATH_WITHIN_24HR).Value = True)

    ' Load cause
    cmbCause.Value = tbl.ListRows(actualRow).Range(1, COL_DEATH_CAUSE).Value

    lblStatus.Caption = "Loaded entry for editing"
    Exit Sub

LoadError:
    MsgBox "Error loading death entry: " & Err.Description, vbCritical, "Load Error"
End Sub

Private Sub PopulateCauses()
    cmbCause.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    Dim causes As Object
    Set causes = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim c As String
        c = Trim(CStr(tbl.ListRows(i).Range(1, COL_DEATH_CAUSE).Value))
        If c <> "" And c <> "0" Then
            If Not causes.Exists(c) Then
                causes.Add c, True
                cmbCause.AddItem c
            End If
        End If
    Next i
End Sub

Private Sub btnSaveNew_Click()
    If SaveDeathEntry() Then
        txtFolderNum.Value = ""
        txtName.Value = ""
        txtAge.Value = ""
        cmbCause.Value = ""
        chkWithin24.Value = False
        UpdatePendingList
        txtFolderNum.SetFocus
    End If
End Sub

Private Sub btnSaveClose_Click()
    If SaveDeathEntry() Then
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    editingRowIndex = 0  ' Clear edit mode
    Unload Me
End Sub

Private Function SaveDeathEntry() As Boolean
    SaveDeathEntry = False

    If cmbWard.ListIndex < 0 Then
        MsgBox "Please select a ward.", vbExclamation
        Exit Function
    End If

    Dim deathDate As Variant
    Dim errMsg As String

    ' Parse and validate date using centralized date utils
    deathDate = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(deathDate) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Function
    End If

    If Not modDateUtils.ValidateDate(deathDate, errMsg) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Function
    End If

    If Trim(txtAge.Value) = "" Or Not IsNumeric(txtAge.Value) Then
        MsgBox "Please enter a valid numeric age.", vbExclamation
        Exit Function
    End If

    Dim ageVal As Long
    ageVal = CLng(txtAge.Value)
    Dim unitVal As String
    unitVal = cmbAgeUnit.Value
    
    ' Validate Age anomalies
    If unitVal = "Years" And (ageVal > 110 Or ageVal = 0) Then
        If MsgBox("Are you sure the age is " & ageVal & " years? This seems unusual. Do you want to review your entry?", vbYesNo + vbExclamation, "Review Age") = vbYes Then
            txtAge.SetFocus
            Exit Function
        End If
    ElseIf unitVal = "Months" And ageVal > 24 Then
        If MsgBox("Are you sure the age is " & ageVal & " months? Usually ages over 24 months are entered in years. Do you want to review your entry?", vbYesNo + vbExclamation, "Review Age") = vbYes Then
            txtAge.SetFocus
            Exit Function
        End If
    ElseIf unitVal = "Days" And ageVal > 28 Then
        If MsgBox("Are you sure the age is " & ageVal & " days? Usually ages over 28 days are entered in months. Do you want to review your entry?", vbYesNo + vbExclamation, "Review Age") = vbYes Then
            txtAge.SetFocus
            Exit Function
        End If
    End If

    Dim sex As String
    If optMale.Value Then sex = "M" Else sex = "F"

    Dim nhis As String
    If optInsured.Value Then nhis = "Insured" Else nhis = "Non-Insured"

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    ' Check if we're editing or creating new
    If editingRowIndex > 0 Then
        ' Edit mode: Update existing row
        UpdateDeathRow editingRowIndex, deathDate, wc, Trim(txtFolderNum.Value), _
            Trim(txtName.Value), CLng(txtAge.Value), cmbAgeUnit.Value, sex, nhis, _
            Trim(cmbCause.Value), chkWithin24.Value
        editingRowIndex = 0  ' Clear edit mode after save
        lblStatus.Caption = "Updated: " & txtName.Value
    Else
        ' New entry mode: Create new row
        SaveDeath deathDate, wc, Trim(txtFolderNum.Value), _
            Trim(txtName.Value), CLng(txtAge.Value), _
            cmbAgeUnit.Value, sex, nhis, _
            Trim(cmbCause.Value), chkWithin24.Value
        lblStatus.Caption = "Saved: " & txtName.Value
    End If

    lblStatus.ForeColor = RGB(0, 128, 0)
    AutoSaveWorkbook
    SaveDeathEntry = True
End Function

Private Sub UpdateDeathRow(rowIndex As Long, deathDate As Variant, wardCode As String, _
    folderNum As String, deceasedName As String, _
    age As Long, ageUnit As String, sex As String, nhis As String, _
    causeOfDeath As String, within24 As Boolean)
    ' Update existing row instead of creating new one
    On Error GoTo UpdateError

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then
        MsgBox "Error: Invalid row index", vbCritical, "Update Error"
        Exit Sub
    End If

    Dim targetRow As ListRow
    Set targetRow = tbl.ListRows(rowIndex)

    With targetRow.Range
        ' Update all fields (keep existing ID, update other fields)
        If IsDate(deathDate) Then
            .Cells(1, COL_DEATH_DATE).Value = CDate(deathDate)
            .Cells(1, COL_DEATH_MONTH).Value = Month(CDate(deathDate))
        Else
            .Cells(1, COL_DEATH_DATE).Value = deathDate
            .Cells(1, COL_DEATH_MONTH).Value = 0
        End If
        .Cells(1, COL_DEATH_WARD_CODE).Value = wardCode
        .Cells(1, COL_DEATH_FOLDER_NUM).Value = folderNum
        .Cells(1, COL_DEATH_NAME).Value = deceasedName
        .Cells(1, COL_DEATH_AGE).Value = age
        .Cells(1, COL_DEATH_AGE_UNIT).Value = ageUnit
        .Cells(1, COL_DEATH_SEX).Value = sex
        .Cells(1, COL_DEATH_NHIS).Value = nhis
        .Cells(1, COL_DEATH_CAUSE).Value = causeOfDeath
        .Cells(1, COL_DEATH_WITHIN_24HR).Value = within24
        .Cells(1, COL_DEATH_TIMESTAMP).Value = Now
        .Cells(1, COL_DEATH_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm"
    End With

    UpdatePendingList
    Exit Sub

UpdateError:
    MsgBox "Error updating death record: " & Err.Description, vbCritical, "Update Error"
End Sub

' ParseDateDth function removed - now using modDateUtils.ParseDate instead

'==============================================================================
' Calendar Picker Button Click Handler
'==============================================================================
Private Sub txtDate_picker_Click()
    ' Show calendar picker and update date field
    modDateUtils.ShowDatePicker txtDate
End Sub
