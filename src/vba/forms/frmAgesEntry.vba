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
    txtDate.Value = Format(Date, "dd/mm/yyyy")

    ' Age Units
    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0 ' Default Years

    ' Defaults
    optMale.Value = True
    optInsured.Value = True

    lblStatus.Caption = "Ready"
    txtAge.SetFocus
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
                tbl.ListRows(i).Range(1, 6).Value & " | " & _
                tbl.ListRows(i).Range(1, 7).Value & " " & _
                tbl.ListRows(i).Range(1, 8).Value & " | " & _
                tbl.ListRows(i).Range(1, 9).Value & " | " & _
                tbl.ListRows(i).Range(1, 10).Value
        End If
    Next i
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub
    On Error GoTo DateError

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
    Dim entryDate As Date
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
    wc = tbl.ListRows(actualRow).Range(1, 6).Value
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
    Dim dateStr As String
    dateStr = Trim(txtDate.Value)
    
    If Not IsDate(dateStr) Then
        MsgBox "Please enter a valid date (e.g. 15/02/2026).", vbExclamation
        txtDate.SetFocus
        Exit Sub
    End If
    
    Dim dt As Date
    dt = CDate(dateStr)

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
        Application.Run "SaveAdmission", dt, wc, "-", "Age Entry", age, unit, sex, nhis
        lblStatus.Caption = "Saved: " & age & " " & unit & " (" & sex & ", " & nhis & ")"
    End If

    ' Post-Save Reset
    lblStatus.ForeColor = RGB(0, 128, 0) ' Green

    txtAge.Value = ""
    cmbAgeUnit.ListIndex = 0 ' Reset to Years
    ' Keep persistent selections (Ward, Date, Sex, NHIS)

    UpdateRecentList
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
        .Cells(1, COL_ADM_PATIENT_NAME).Value = "Age Entry"  ' Ages entry uses this for patient name
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
    MsgBox "Error updating age entry: " & Err.Description, vbCritical, "Update Error"
End Sub

Private Sub btnClose_Click()
    editingRowIndex = 0  ' Clear edit mode
    Unload Me
End Sub
