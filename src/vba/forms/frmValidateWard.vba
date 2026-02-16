Option Explicit

Private wardCodes As Variant
Private wardNames As Variant

'==============================================================================
' Form Initialization
'==============================================================================
Private Sub UserForm_Initialize()
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Load months
    cmbMonth.Clear
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
    cmbMonth.ListIndex = Month(Date) - 1  ' Current month

    ' Load wards
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Setup list box headers
    lstResults.Clear
    lstResults.ColumnCount = 4
    lstResults.ColumnWidths = "80;60;80;70"

    ' Add header row to list box
    lstResults.AddItem "Date" & vbTab & "Daily" & vbTab & "Individual" & vbTab & "Status"

    lblSummary.Caption = "Select month and ward, then click 'Validate Month'"
    lblSummary.ForeColor = RGB(100, 100, 100)
End Sub

'==============================================================================
' Validate Button Click Handler
'==============================================================================
Private Sub btnValidate_Click()
    If cmbMonth.ListIndex < 0 Or cmbWard.ListIndex < 0 Then
        MsgBox "Please select both month and ward", vbExclamation, "Validation Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    lstResults.Clear
    ' Re-add header row
    lstResults.AddItem "Date" & vbTab & "Daily" & vbTab & "Individual" & vbTab & "Status"

    Dim monthIdx As Long
    monthIdx = cmbMonth.ListIndex + 1  ' Convert to 1-based

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim reportYear As Long
    reportYear = GetReportYear()

    ' Get validation report
    Dim results As Variant
    results = GetMonthlyValidationReport(monthIdx, wc, reportYear)

    If IsEmpty(results) Then
        lblSummary.Caption = "No data found for " & cmbMonth.Value & " " & reportYear & " in " & cmbWard.Value
        lblSummary.ForeColor = RGB(128, 128, 128)
        Application.Cursor = xlDefault
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Populate list box
    Dim mismatchCount As Long
    mismatchCount = 0

    Dim i As Long
    For i = 1 To UBound(results, 1)
        lstResults.AddItem Format(results(i, 1), "dd/mm/yyyy") & vbTab & _
                          results(i, 2) & vbTab & _
                          results(i, 3) & vbTab & _
                          results(i, 4)

        If results(i, 4) = "MISMATCH" Then
            mismatchCount = mismatchCount + 1
        End If
    Next i

    ' Update summary
    Dim totalDays As Long
    totalDays = UBound(results, 1)

    If mismatchCount = 0 Then
        lblSummary.Caption = "All OK: " & totalDays & " days checked, no mismatches"
        lblSummary.ForeColor = RGB(0, 128, 0)  ' Green
    Else
        lblSummary.Caption = "WARNING: " & mismatchCount & " mismatch(es) found (out of " & totalDays & " days)"
        lblSummary.ForeColor = RGB(255, 0, 0)  ' Red
    End If

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' Export Button Click Handler
' Exports validation results to a new worksheet for further analysis
'==============================================================================
Private Sub btnExport_Click()
    If lstResults.ListCount <= 1 Then ' Only header row
        MsgBox "No validation results to export. Please run validation first.", vbExclamation, "Export Error"
        Exit Sub
    End If

    On Error GoTo ExportError

    ' Create new worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Validation_" & Format(Now, "yyyymmdd_hhmmss")

    ' Add title
    ws.Range("A1").Value = "Ward Validation Report"
    ws.Range("A2").Value = "Ward: " & cmbWard.Value
    ws.Range("A3").Value = "Month: " & cmbMonth.Value & " " & GetReportYear()
    ws.Range("A4").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm")

    ' Add headers
    ws.Range("A6").Value = "Date"
    ws.Range("B6").Value = "Daily Total"
    ws.Range("C6").Value = "Individual Count"
    ws.Range("D6").Value = "Status"

    ' Add data (skip first row which is header)
    Dim i As Long
    Dim row As Long
    row = 7
    For i = 1 To lstResults.ListCount - 1  ' Skip header row
        Dim itemText As String
        itemText = lstResults.List(i)

        ' Parse tab-delimited data
        Dim parts() As String
        parts = Split(itemText, vbTab)

        If UBound(parts) >= 3 Then
            ws.Cells(row, 1).Value = parts(0)  ' Date
            ws.Cells(row, 2).Value = CLng(parts(1))  ' Daily Total
            ws.Cells(row, 3).Value = CLng(parts(2))  ' Individual Count
            ws.Cells(row, 4).Value = parts(3)  ' Status

            ' Color code mismatches
            If parts(3) = "MISMATCH" Then
                ws.Cells(row, 4).Interior.Color = RGB(255, 200, 200)  ' Light red
            End If

            row = row + 1
        End If
    Next i

    ' Format
    ws.Range("A6:D6").Font.Bold = True
    ws.Columns("A:D").AutoFit

    ' Activate the new sheet
    ws.Activate

    MsgBox "Validation results exported to worksheet: " & ws.Name, vbInformation, "Export Complete"
    Exit Sub

ExportError:
    MsgBox "Error exporting results: " & Err.Description, vbCritical, "Export Error"
End Sub

'==============================================================================
' Close Button Click Handler
'==============================================================================
Private Sub btnClose_Click()
    Unload Me
End Sub
