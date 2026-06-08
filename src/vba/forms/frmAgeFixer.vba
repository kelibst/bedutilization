Option Explicit

Private Sub UserForm_Initialize()
    With lstAnomalies
        .ColumnCount = 9
        ' Table, RowIdx, Ward, Date, Patient, CurAge, CurUnit, SugAge, SugUnit
        .ColumnWidths = "0;0;40;60;80;40;40;40;40"
    End With
    
    cmbNewUnit.AddItem "Years"
    cmbNewUnit.AddItem "Months"
    cmbNewUnit.AddItem "Days"
    
    ScanDatabase
End Sub

Private Sub btnRefresh_Click()
    ScanDatabase
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub ScanDatabase()
    lstAnomalies.Clear
    txtNewAge.Value = ""
    cmbNewUnit.ListIndex = -1
    lblStatus.Caption = "Scanning database..."
    
    Dim tblAdm As ListObject
    Dim tblDth As ListObject
    On Error Resume Next
    Set tblAdm = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")
    Set tblDth = ThisWorkbook.Sheets("Deaths").ListObjects("tblDeaths")
    On Error GoTo 0
    
    Dim anomalyCount As Long
    anomalyCount = 0
    
    ' Scan Admissions (COL_ADM_AGE = 7, COL_ADM_AGE_UNIT = 8, WARD=4, DATE=2, PATIENT=6)
    If Not tblAdm Is Nothing Then
        ScanTable tblAdm, "Adm", 7, 8, 4, 2, 6, anomalyCount
    End If
    
    ' Scan Deaths (COL_DEATH_AGE = 7, COL_DEATH_AGE_UNIT = 8, WARD=4, DATE=2, PATIENT=6)
    If Not tblDth Is Nothing Then
        ScanTable tblDth, "Dth", 7, 8, 4, 2, 6, anomalyCount
    End If
    
    lblStatus.Caption = "Found " & anomalyCount & " anomalies."
End Sub

Private Sub ScanTable(tbl As ListObject, tblType As String, colAge As Long, colUnit As Long, colWard As Long, colDate As Long, colPatient As Long, ByRef anomalyCount As Long)
    Dim i As Long
    Dim vAge As Variant
    Dim vUnit As String
    Dim aAge As Long
    Dim sugAge As Long
    Dim sugUnit As String
    Dim isAnomaly As Boolean
    
    For i = 1 To tbl.ListRows.Count
        vAge = tbl.ListRows(i).Range(1, colAge).Value
        vUnit = Trim(CStr(tbl.ListRows(i).Range(1, colUnit).Value))
        
        If IsNumeric(vAge) And vAge <> "" Then
            aAge = CLng(vAge)
            isAnomaly = False
            sugAge = aAge
            sugUnit = vUnit
            
            If vUnit = "Days" And aAge > 28 Then
                isAnomaly = True
                sugAge = aAge \ 30
                If sugAge > 11 Then
                    sugAge = sugAge \ 12
                    sugUnit = "Years"
                Else
                    sugUnit = "Months"
                End If
                If sugAge = 0 Then sugAge = 1 ' Fallback
            ElseIf vUnit = "Months" And aAge > 11 Then
                isAnomaly = True
                sugAge = aAge \ 12
                If sugAge = 0 Then sugAge = 1 ' Minimum 1 year if it was like 12 months but integer div is 1 anyway
                sugUnit = "Years"
            ElseIf vUnit = "Years" And aAge = 0 Then
                isAnomaly = True
                sugAge = 0 ' User requested 0 days or blank, we will leave it 0
                sugUnit = "Days" ' Default suggestion 0 days
            End If
            
            If isAnomaly Then
                lstAnomalies.AddItem tblType
                lstAnomalies.List(lstAnomalies.ListCount - 1, 1) = i
                lstAnomalies.List(lstAnomalies.ListCount - 1, 2) = tbl.ListRows(i).Range(1, colWard).Value
                
                Dim dtVal As Variant
                dtVal = tbl.ListRows(i).Range(1, colDate).Value
                If IsDate(dtVal) Then
                    lstAnomalies.List(lstAnomalies.ListCount - 1, 3) = Format(dtVal, "dd/mm/yyyy")
                Else
                    lstAnomalies.List(lstAnomalies.ListCount - 1, 3) = dtVal
                End If
                
                lstAnomalies.List(lstAnomalies.ListCount - 1, 4) = tbl.ListRows(i).Range(1, colPatient).Value
                lstAnomalies.List(lstAnomalies.ListCount - 1, 5) = aAge
                lstAnomalies.List(lstAnomalies.ListCount - 1, 6) = vUnit
                lstAnomalies.List(lstAnomalies.ListCount - 1, 7) = sugAge
                lstAnomalies.List(lstAnomalies.ListCount - 1, 8) = sugUnit
                
                anomalyCount = anomalyCount + 1
            End If
        End If
    Next i
End Sub

Private Sub lstAnomalies_Click()
    If lstAnomalies.ListIndex >= 0 Then
        txtNewAge.Value = lstAnomalies.List(lstAnomalies.ListIndex, 7)
        cmbNewUnit.Value = lstAnomalies.List(lstAnomalies.ListIndex, 8)
    End If
End Sub

Private Sub btnApply_Click()
    If lstAnomalies.ListIndex < 0 Then
        MsgBox "Please select an anomaly from the list.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtNewAge.Value) = "" Or Not IsNumeric(txtNewAge.Value) Then
        MsgBox "Please enter a valid numeric age.", vbExclamation
        Exit Sub
    End If
    
    If cmbNewUnit.ListIndex < 0 Then
        MsgBox "Please select an age unit.", vbExclamation
        Exit Sub
    End If
    
    Dim tblType As String
    Dim rowIdx As Long
    Dim newAge As Long
    Dim newUnit As String
    
    tblType = lstAnomalies.List(lstAnomalies.ListIndex, 0)
    rowIdx = CLng(lstAnomalies.List(lstAnomalies.ListIndex, 1))
    newAge = CLng(txtNewAge.Value)
    newUnit = cmbNewUnit.Value
    
    Dim tbl As ListObject
    If tblType = "Adm" Then
        Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")
        tbl.ListRows(rowIdx).Range(1, 7).Value = newAge
        tbl.ListRows(rowIdx).Range(1, 8).Value = newUnit
    Else
        Set tbl = ThisWorkbook.Sheets("Deaths").ListObjects("tblDeaths")
        tbl.ListRows(rowIdx).Range(1, 7).Value = newAge
        tbl.ListRows(rowIdx).Range(1, 8).Value = newUnit
    End If
    
    lblStatus.Caption = "Applied update successfully."
    lblStatus.ForeColor = RGB(0, 128, 0)
    
    ' Remove from listbox
    lstAnomalies.RemoveItem lstAnomalies.ListIndex
    txtNewAge.Value = ""
    cmbNewUnit.ListIndex = -1
    
    ' Decrease count
    Dim curCount As Long
    curCount = 0
    On Error Resume Next
    curCount = CLng(Split(lblStatus.Caption, " ")(1))
    On Error GoTo 0
    
    AutoSaveWorkbook
End Sub
