Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private editingRowIndex As Long  ' 0 = new entry, >0 = editing specific row

Private Sub UserForm_Initialize()
    editingRowIndex = 0  ' Start in new entry mode
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Load age units FIRST (before ward selection triggers cmbWard_Change)
    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0

    ' Now load wards (setting ListIndex will fire cmbWard_Change safely)
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    txtDate.Value = Format(Date, "dd/mm/yyyy")
    txtAge.Value = ""
    txtPatientID.Value = ""
    txtPatientName.Value = ""
    optMale.Value = True
    optInsured.Value = True
    UpdateRecentList
End Sub

Private Sub UpdateRecentList()
    lstRecent.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim i As Long
    For i = startRow To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) And _
           tbl.ListRows(i).Range(1, 2).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 2).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 4).Value & " | " & _
                tbl.ListRows(i).Range(1, 6).Value & " | Age: " & _
                tbl.ListRows(i).Range(1, 7).Value
        End If
    Next i
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

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
    txtPatientID.Value = tbl.ListRows(actualRow).Range(1, 5).Value
    txtPatientName.Value = tbl.ListRows(actualRow).Range(1, 6).Value
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
End Sub

Private Sub cmbWard_Change()
    ' Auto-set age unit based on ward
    If cmbWard.ListIndex >= 0 Then
        Dim wc As String
        wc = wardCodes(cmbWard.ListIndex)
        If wc = "NICU" Then
            cmbAgeUnit.ListIndex = 2  ' Days
        ElseIf wc = "CW" Then
            cmbAgeUnit.ListIndex = 0  ' Years (but user can change)
        Else
            cmbAgeUnit.ListIndex = 0  ' Years
        End If
    End If
End Sub

Private Sub btnSaveNew_Click()
    If SaveAdmissionEntry() Then
        ' Clear for next entry but keep date and ward
        txtPatientID.Value = ""
        txtPatientName.Value = ""
        txtAge.Value = ""
        txtPatientID.SetFocus
        UpdateRecentList
    End If
End Sub

Private Sub btnSaveClose_Click()
    If SaveAdmissionEntry() Then
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    editingRowIndex = 0  ' Clear edit mode
    Unload Me
End Sub

Private Function SaveAdmissionEntry() As Boolean
    SaveAdmissionEntry = False

    If cmbWard.ListIndex < 0 Then
        MsgBox "Please select a ward.", vbExclamation
        Exit Function
    End If

    Dim admDate As Variant
    admDate = Trim(txtDate.Value)
    
    If Not IsDate(admDate) Then
        MsgBox "Please enter a valid date (e.g., 14/02/2026).", vbExclamation
        txtDate.SetFocus
        Exit Function
    End If
    
    ' Convert to real date to ensure consistency
    admDate = CDate(admDate)

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
        UpdateAdmissionRow editingRowIndex, admDate, wc, Trim(txtPatientID.Value), _
            Trim(txtPatientName.Value), CLng(txtAge.Value), cmbAgeUnit.Value, sex, nhis
        editingRowIndex = 0  ' Clear edit mode after save
        lblStatus.Caption = "Updated: " & txtPatientName.Value
    Else
        ' New entry mode: Create new row
        SaveAdmission admDate, wc, Trim(txtPatientID.Value), _
            Trim(txtPatientName.Value), CLng(txtAge.Value), _
            cmbAgeUnit.Value, sex, nhis
        lblStatus.Caption = "Saved: " & txtPatientName.Value
    End If

    lblStatus.ForeColor = RGB(0, 128, 0)
    SaveAdmissionEntry = True
End Function

Private Sub UpdateAdmissionRow(rowIndex As Long, admDate As Variant, wardCode As String, _
    patientID As String, patientName As String, _
    age As Long, ageUnit As String, sex As String, nhis As String)
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
        .Cells(1, COL_ADM_DATE).Value = admDate
        .Cells(1, COL_ADM_DATE).NumberFormat = "yyyy-mm-dd"
        .Cells(1, COL_ADM_MONTH).Value = Month(admDate)
        .Cells(1, COL_ADM_WARD_CODE).Value = wardCode
        .Cells(1, COL_ADM_PATIENT_ID).Value = patientID
        .Cells(1, COL_ADM_PATIENT_NAME).Value = patientName
        .Cells(1, COL_ADM_AGE).Value = age
        .Cells(1, COL_ADM_AGE_UNIT).Value = ageUnit
        .Cells(1, COL_ADM_SEX).Value = sex
        .Cells(1, COL_ADM_NHIS).Value = nhis
        .Cells(1, COL_ADM_TIMESTAMP).Value = Now
        .Cells(1, COL_ADM_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm"
    End With

    UpdateRecentList
    Exit Sub

UpdateError:
    MsgBox "Error updating entry: " & Err.Description, vbCritical, "Update Error"
End Sub

Private Function ParseDateAdm(dateStr As String) As Date
    On Error GoTo badDate

    ' Validate input
    If Trim(dateStr) = "" Then
        MsgBox "Date field is empty. Please enter a valid date.", vbExclamation, "Invalid Date"
        ParseDateAdm = #1/1/1900#
        Exit Function
    End If

    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        ParseDateAdm = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        Exit Function
    End If

    ParseDateAdm = CDate(dateStr)
    Exit Function

badDate:
    MsgBox "Invalid date format: " & dateStr & vbCrLf & _
           "Please use dd/mm/yyyy format (e.g., 13/02/2026)", _
           vbExclamation, "Invalid Date"
    ParseDateAdm = #1/1/1900#
End Function
