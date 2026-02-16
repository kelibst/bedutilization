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

    txtDate.Value = modDateUtils.FormatDateDisplay(Date)
    txtAge.Value = ""
    txtPatientID.Value = ""
    txtPatientName.Value = ""
    optMale.Value = True
    optInsured.Value = True

    ' Initialize date filter controls
    On Error Resume Next
    ' Handle both DTPicker and TextBox
    If TypeName(Me.Controls("dtpFilterDate")) = "DTPicker" Then
        dtpFilterDate.Value = Date
    Else
        dtpFilterDate.Value = Format(Date, "dd/mm/yyyy")
    End If
    On Error GoTo 0

    UpdateRecentList
    UpdateValidationDisplay
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
                        tbl.ListRows(i).Range(1, 4).Value & " | " & _
                        tbl.ListRows(i).Range(1, 6).Value & " | Age: " & _
                        tbl.ListRows(i).Range(1, 7).Value
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
                    tbl.ListRows(i).Range(1, 4).Value & " | " & _
                    tbl.ListRows(i).Range(1, 6).Value & " | Age: " & _
                    tbl.ListRows(i).Range(1, 7).Value
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
        lblValidation.Caption = "Daily Total: " & dailyTotal & " | Individual Count: " & individualCount & " [OK]"
        lblValidation.ForeColor = RGB(0, 128, 0)  ' Green
    Else
        If dailyTotal = 0 And InStr(validationMsg, "No daily bed-state") > 0 Then
            lblValidation.Caption = "Daily Total: Not entered yet"
            lblValidation.ForeColor = RGB(128, 128, 128)  ' Gray
        Else
            lblValidation.Caption = "Daily Total: " & dailyTotal & " | Individual Count: " & individualCount & " [MISMATCH]"
            lblValidation.ForeColor = RGB(255, 0, 0)  ' Red
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

    ' Parse the selected item to extract admission date, ward, and patient name
    Dim selectedItem As String
    selectedItem = lstRecent.List(lstRecent.ListIndex)

    ' Format: "dd/mm/yyyy | Ward | PatientName | Age: ##"
    Dim datePart As String, wardPart As String, namePart As String
    Dim firstPipe As Integer, secondPipe As Integer, thirdPipe As Integer
    firstPipe = InStr(selectedItem, "|")
    secondPipe = InStr(firstPipe + 1, selectedItem, "|")
    thirdPipe = InStr(secondPipe + 1, selectedItem, "|")

    datePart = Trim(Left(selectedItem, firstPipe - 1))
    wardPart = Trim(Mid(selectedItem, firstPipe + 1, secondPipe - firstPipe - 1))
    namePart = Trim(Mid(selectedItem, secondPipe + 1, thirdPipe - secondPipe - 1))

    ' Find the matching row in the table by date, ward, and patient name
    Dim entryDate As Date
    entryDate = CDate(datePart)

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim actualRow As Long
    actualRow = 0
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) Then
            Dim checkDate As Date
            checkDate = CDate(tbl.ListRows(i).Range(1, 2).Value)
            Dim checkWard As String
            checkWard = tbl.ListRows(i).Range(1, 4).Value
            Dim checkName As String
            checkName = tbl.ListRows(i).Range(1, 6).Value

            If DateValue(checkDate) = DateValue(entryDate) And _
               checkWard = wardPart And checkName = namePart Then
                actualRow = i
                Exit For
            End If
        End If
    Next i

    If actualRow = 0 Then
        MsgBox "Could not find the selected entry in the table.", vbExclamation
        Exit Sub
    End If

    ' Load the record
    LoadRecordFromRow actualRow
    lblStatus.Caption = "Editing: " & namePart
    lblStatus.ForeColor = RGB(255, 128, 0)  ' Orange
    Exit Sub

LoadError:
    MsgBox "Error loading admission entry: " & Err.Description, vbCritical, "Load Error"
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
    UpdateValidationDisplay
End Sub

Private Sub btnSaveNew_Click()
    If SaveAdmissionEntry() Then
        ' Clear for next entry but keep date and ward
        txtPatientID.Value = ""
        txtPatientName.Value = ""
        txtAge.Value = ""
        txtPatientID.SetFocus
        UpdateRecentList
        UpdateValidationDisplay
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
    Dim errMsg As String

    ' Parse and validate date using centralized date utils
    admDate = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(admDate) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Function
    End If

    If Not modDateUtils.ValidateDate(admDate, errMsg) Then
        MsgBox errMsg, vbExclamation, "Invalid Date"
        txtDate.SetFocus
        Exit Function
    End If

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

' ParseDateAdm function removed - now using modDateUtils.ParseDate instead

'==============================================================================
' Calendar Picker Button Click Handler
'==============================================================================
Private Sub txtDate_picker_Click()
    ' Show calendar picker and update date field
    If modDateUtils.ShowDatePicker(txtDate) Then
        ' Date was updated, check for existing record
        CheckAndLoadExistingRecord
    End If
End Sub

'==============================================================================
' Date Field Change Handler
' When date changes (typed or picked), check for existing records
'==============================================================================
Private Sub txtDate_AfterUpdate()
    CheckAndLoadExistingRecord
    UpdateValidationDisplay
End Sub

'==============================================================================
' Check if record exists for current date + ward, and load it
'==============================================================================
Private Sub CheckAndLoadExistingRecord()
    ' Only check if we have a valid date and ward selected
    If Trim(txtDate.Value) = "" Or cmbWard.ListIndex < 0 Then Exit Sub

    On Error Resume Next
    Dim checkDate As Variant
    Dim errMsg As String
    checkDate = modDateUtils.ParseDate(txtDate.Value, errMsg)
    If IsEmpty(checkDate) Then Exit Sub

    ' Get current ward code
    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    ' Search for existing record with this date + ward
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim i As Long
    For i = tbl.ListRows.Count To 1 Step -1  ' Search backwards (most recent first)
        If Not IsEmpty(tbl.ListRows(i).Range(1, 2).Value) Then
            Dim entryDate As Date
            entryDate = CDate(tbl.ListRows(i).Range(1, 2).Value)
            Dim entryWard As String
            entryWard = tbl.ListRows(i).Range(1, 4).Value

            ' If we find a match, load it
            If DateValue(entryDate) = DateValue(checkDate) And entryWard = wc Then
                LoadRecordFromRow i
                lblStatus.Caption = "Editing existing record for " & Format(checkDate, "dd/mm/yyyy")
                lblStatus.ForeColor = RGB(255, 128, 0)  ' Orange
                Exit Sub
            End If
        End If
    Next i

    ' No existing record found - stay in new entry mode
    editingRowIndex = 0
    lblStatus.Caption = "New admission for " & Format(checkDate, "dd/mm/yyyy")
    lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
    UpdateValidationDisplay
End Sub

'==============================================================================
' Load a specific record by row index
'==============================================================================
Private Sub LoadRecordFromRow(rowIndex As Long)
    On Error GoTo LoadError

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    ' Store the row we're editing
    editingRowIndex = rowIndex

    ' Load the record (date is already set)
    txtDate.Value = Format(tbl.ListRows(rowIndex).Range(1, 2).Value, "dd/mm/yyyy")

    ' Load ward
    Dim wc As String
    wc = tbl.ListRows(rowIndex).Range(1, 4).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = wc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load patient details
    txtPatientID.Value = tbl.ListRows(rowIndex).Range(1, 5).Value
    txtPatientName.Value = tbl.ListRows(rowIndex).Range(1, 6).Value
    txtAge.Value = CStr(tbl.ListRows(rowIndex).Range(1, 7).Value)
    cmbAgeUnit.Value = tbl.ListRows(rowIndex).Range(1, 8).Value

    ' Load sex
    If tbl.ListRows(rowIndex).Range(1, 9).Value = "M" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If

    ' Load NHIS
    If tbl.ListRows(rowIndex).Range(1, 10).Value = "Insured" Then
        optInsured.Value = True
    Else
        optNonInsured.Value = True
    End If
    Exit Sub

LoadError:
    MsgBox "Error loading record: " & Err.Description, vbCritical
End Sub
