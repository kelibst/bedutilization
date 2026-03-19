Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private allResults As Variant  ' Full results from GetMonthlyValidationReport (all days)

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

    ' Setup list box: Date | Daily | Individual | Delta | Status
    lstResults.Clear
    lstResults.ColumnCount = 5
    lstResults.ColumnWidths = "58;46;68;46;64"

    ' Add column header row
    lstResults.AddItem "Date" & vbTab & "Daily" & vbTab & "Individual" & vbTab & "Delta" & vbTab & "Status"

    lblSummary.Caption = "Select month and ward, then click Validate Month"
    lblSummary.ForeColor = RGB(100, 100, 100)

    lblErrorDates.Caption = ""
    lblErrorDates.Visible = False
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

    Dim monthIdx As Long
    monthIdx = cmbMonth.ListIndex + 1  ' Convert to 1-based

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim reportYear As Long
    reportYear = GetReportYear()

    ' Get full monthly report (all days)
    allResults = GetMonthlyValidationReport(monthIdx, wc, reportYear)

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    If IsEmpty(allResults) Then
        allResults = Empty
        lblSummary.Caption = "No data found for " & cmbMonth.Value & " " & reportYear & " - " & cmbWard.Value
        lblSummary.ForeColor = RGB(128, 128, 128)
        lblErrorDates.Visible = False
        lstResults.Clear
        lstResults.AddItem "Date" & vbTab & "Daily" & vbTab & "Individual" & vbTab & "Delta" & vbTab & "Status"
        Exit Sub
    End If

    RefreshDisplay
End Sub

'==============================================================================
' Populate / Refresh the ListBox (respects Errors Only checkbox)
'==============================================================================
Private Sub RefreshDisplay()
    lstResults.Clear
    lstResults.AddItem "Date" & vbTab & "Daily" & vbTab & "Individual" & vbTab & "Delta" & vbTab & "Status"

    If IsEmpty(allResults) Then Exit Sub

    Dim okCount As Long, mismatchCount As Long, noEntryCount As Long
    GetValidationSummary allResults, okCount, mismatchCount, noEntryCount

    ' Build mismatch date list for the error label
    Dim errorDateList As String
    errorDateList = ""

    Dim showErrorsOnly As Boolean
    showErrorsOnly = chkErrorsOnly.Value

    Dim i As Long
    For i = 1 To UBound(allResults, 1)
        Dim status As String
        status = allResults(i, 4)

        ' Accumulate mismatch dates regardless of filter
        If status = "MISMATCH" Then
            If Len(errorDateList) > 0 Then errorDateList = errorDateList & ", "
            errorDateList = errorDateList & Format(allResults(i, 1), "dd/mm")
        End If

        ' Apply filter
        Dim showRow As Boolean
        If showErrorsOnly Then
            showRow = (status = "MISMATCH")
        Else
            showRow = True
        End If

        If showRow Then
            Dim dateStr As String
            Dim dailyStr As String
            Dim indivStr As String
            Dim deltaStr As String
            Dim statusStr As String

            dateStr = Format(allResults(i, 1), "dd/mm/yyyy")

            If status = "NO ENTRY" Then
                dailyStr = "--"
                indivStr = "--"
                deltaStr = "--"
                statusStr = "NO ENTRY"
            Else
                dailyStr = CStr(allResults(i, 2))
                indivStr = CStr(allResults(i, 3))
                Dim delta As Long
                delta = CLng(allResults(i, 5))
                deltaStr = IIf(delta = 0, "0", IIf(delta > 0, "+" & CStr(delta), CStr(delta)))
                statusStr = status
            End If

            lstResults.AddItem dateStr & vbTab & dailyStr & vbTab & indivStr & vbTab & deltaStr & vbTab & statusStr
        End If
    Next i

    ' Update summary label
    Dim totalDays As Long
    totalDays = UBound(allResults, 1)

    If mismatchCount = 0 Then
        lblSummary.Caption = "All OK  |  " & okCount & " days matched  |  " & noEntryCount & " no entry  |  " & totalDays & " days total"
        lblSummary.ForeColor = RGB(0, 128, 0)
    Else
        lblSummary.Caption = mismatchCount & " mismatch(es)  |  " & okCount & " OK  |  " & noEntryCount & " no entry  |  " & totalDays & " days total"
        lblSummary.ForeColor = RGB(192, 0, 0)
    End If

    ' Show mismatch date list
    If Len(errorDateList) > 0 Then
        lblErrorDates.Caption = "Mismatch dates:  " & errorDateList
        lblErrorDates.Visible = True
    Else
        lblErrorDates.Caption = ""
        lblErrorDates.Visible = False
    End If
End Sub

'==============================================================================
' Errors Only checkbox: re-filter the list without re-querying data
'==============================================================================
Private Sub chkErrorsOnly_Click()
    If Not IsEmpty(allResults) Then
        RefreshDisplay
    End If
End Sub

'==============================================================================
' Export Button Click Handler
'==============================================================================
Private Sub btnExport_Click()
    If IsEmpty(allResults) Then
        MsgBox "No validation results to export. Please run validation first.", vbExclamation, "Export Error"
        Exit Sub
    End If

    On Error GoTo ExportError

    ' Create new worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Validation_" & Format(Now, "yyyymmdd_hhmmss")

    ' Title block
    ws.Range("A1").Value = "Ward Validation Report"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Ward: " & cmbWard.Value
    ws.Range("A3").Value = "Month: " & cmbMonth.Value & " " & GetReportYear()
    ws.Range("A4").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm")

    ' Column headers
    ws.Range("A6").Value = "Date"
    ws.Range("B6").Value = "Daily Total"
    ws.Range("C6").Value = "Individual Count"
    ws.Range("D6").Value = "Delta"
    ws.Range("E6").Value = "Status"
    ws.Range("A6:E6").Font.Bold = True

    ' Data rows - export ALL results regardless of current filter
    Dim row As Long
    row = 7

    Dim i As Long
    For i = 1 To UBound(allResults, 1)
        Dim status As String
        status = allResults(i, 4)

        ws.Cells(row, 1).Value = Format(allResults(i, 1), "dd/mm/yyyy")

        If status = "NO ENTRY" Then
            ws.Cells(row, 2).Value = ""
            ws.Cells(row, 3).Value = ""
            ws.Cells(row, 4).Value = ""
            ws.Cells(row, 5).Value = "NO ENTRY"
            ws.Cells(row, 5).Font.Color = RGB(128, 128, 128)
        Else
            ws.Cells(row, 2).Value = CLng(allResults(i, 2))
            ws.Cells(row, 3).Value = CLng(allResults(i, 3))
            ws.Cells(row, 4).Value = CLng(allResults(i, 5))
            ws.Cells(row, 5).Value = status

            If status = "MISMATCH" Then
                ws.Cells(row, 5).Interior.Color = RGB(255, 180, 180)
                ws.Cells(row, 5).Font.Bold = True
                ws.Cells(row, 4).Interior.Color = RGB(255, 220, 180)
            End If
        End If

        row = row + 1
    Next i

    ' Summary below data
    row = row + 1
    Dim okCount As Long, mismatchCount As Long, noEntryCount As Long
    GetValidationSummary allResults, okCount, mismatchCount, noEntryCount

    ws.Cells(row, 1).Value = "Summary"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "OK days:": ws.Cells(row, 2).Value = okCount
    row = row + 1
    ws.Cells(row, 1).Value = "Mismatches:": ws.Cells(row, 2).Value = mismatchCount
    row = row + 1
    ws.Cells(row, 1).Value = "No entry:": ws.Cells(row, 2).Value = noEntryCount

    ws.Columns("A:E").AutoFit

    ws.Activate
    MsgBox "Validation results exported to worksheet: " & ws.Name, vbInformation, "Export Complete"
    Exit Sub

ExportError:
    MsgBox "Error exporting results: " & Err.Description, vbCritical, "Export Error"
End Sub

'==============================================================================
' Close Button
'==============================================================================
Private Sub btnClose_Click()
    Unload Me
End Sub
