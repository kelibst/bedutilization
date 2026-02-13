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

    txtDate.Value = Format(Date, "dd/mm/yyyy")
    optMale.Value = True
    optInsured.Value = True
    chkWithin24.Value = False

    ' Populate cause of death combo with previous entries
    PopulateCauses
    UpdateRecentList
End Sub

Private Sub UpdateRecentList()
    lstRecent.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim i As Long
    For i = startRow To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) And _
           tbl.ListRows(i).Range(1, 2).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 2).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 6).Value & " | " & _
                tbl.ListRows(i).Range(1, 5).Value & " | " & _
                tbl.ListRows(i).Range(1, 4).Value
        End If
    Next i
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    ' Calculate actual row (last 10 entries)
    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim actualRow As Long
    actualRow = startRow + lstRecent.ListIndex

    If actualRow > tbl.ListRows.Count Then Exit Sub

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
    chkWithin24.Value = (tbl.ListRows(actualRow).Range(1, 11).Value = True)

    ' Load cause
    cmbCause.Value = tbl.ListRows(actualRow).Range(1, 10).Value

    lblStatus.Caption = "Loaded entry for editing"
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
        c = Trim(CStr(tbl.ListRows(i).Range(1, 11).Value))
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
        UpdateRecentList
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
    deathDate = Trim(txtDate.Value)
    
    If Not IsDate(deathDate) Then
        MsgBox "Please enter a valid date (e.g., 14/02/2026).", vbExclamation
        txtDate.SetFocus
        Exit Function
    End If
    
    deathDate = CDate(deathDate)

    If Trim(txtAge.Value) = "" Or Not IsNumeric(txtAge.Value) Then
        MsgBox "Please enter a valid age.", vbExclamation
        Exit Function
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

    UpdateRecentList
    Exit Sub

UpdateError:
    MsgBox "Error updating death record: " & Err.Description, vbCritical, "Update Error"
End Sub

Private Function ParseDateDth(dateStr As String) As Date
    On Error GoTo badDate

    ' Validate input
    If Trim(dateStr) = "" Then
        MsgBox "Date field is empty. Please enter a valid date.", vbExclamation, "Invalid Date"
        ParseDateDth = #1/1/1900#
        Exit Function
    End If

    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        ParseDateDth = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        Exit Function
    End If

    ParseDateDth = CDate(dateStr)
    Exit Function

badDate:
    MsgBox "Invalid date format: " & dateStr & vbCrLf & _
           "Please use dd/mm/yyyy format (e.g., 13/02/2026)", _
           vbExclamation, "Invalid Date"
    ParseDateDth = #1/1/1900#
End Function
