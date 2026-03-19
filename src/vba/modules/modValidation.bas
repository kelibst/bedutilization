'###################################################################
'# MODULE: modValidation
'# PURPOSE: Admission validation logic - verify individual admissions
'#          match daily bed-state totals
'###################################################################

Option Explicit

'===================================================================
' COUNT INDIVIDUAL ADMISSIONS
'===================================================================
Public Function CountIndividualAdmissions(entryDate As Date, wardCode As String) As Long
    ' Count individual admission records for a specific date and ward
    ' Includes both named patients and "Age Entry" bulk entries
    '
    ' Args:
    '   entryDate: The admission date to check
    '   wardCode: The ward code (e.g., "MW", "FW")
    '
    ' Returns:
    '   Long: Count of admission records matching date/ward

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim rowDate As Variant
        rowDate = tbl.ListRows(i).Range(1, COL_ADM_DATE).Value

        If IsDate(rowDate) Then
            Dim rowWard As String
            rowWard = Trim(CStr(tbl.ListRows(i).Range(1, COL_ADM_WARD_CODE).Value))

            ' Compare dates using DateValue to ignore time component
            If DateValue(CDate(rowDate)) = DateValue(entryDate) And _
               rowWard = wardCode Then
                count = count + 1
            End If
        End If
    Next i

    CountIndividualAdmissions = count
    Exit Function

ErrorHandler:
    ' On error, return 0 (safer than propagating error)
    CountIndividualAdmissions = 0
End Function

'===================================================================
' GET DAILY ADMISSION TOTAL
'===================================================================
Public Function GetDailyAdmissionTotal(entryDate As Date, wardCode As String) As Variant
    ' Get the admission count from tblDaily for a specific date/ward
    '
    ' Args:
    '   entryDate: The date to check
    '   wardCode: The ward code
    '
    ' Returns:
    '   Long: Admission count if found
    '   Empty: If no daily entry exists

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim i As Long
    For i = 1 To tbl.ListRows.count
        Dim rowDate As Variant
        rowDate = tbl.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value

        If IsDate(rowDate) Then
            Dim rowWard As String
            rowWard = Trim(CStr(tbl.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value))

            ' Compare dates using DateValue to ignore time component
            If DateValue(CDate(rowDate)) = DateValue(entryDate) And _
               rowWard = wardCode Then
                ' Found matching entry - return admission count
                GetDailyAdmissionTotal = CLng(tbl.ListRows(i).Range(1, COL_DAILY_ADMISSIONS).Value)
                Exit Function
            End If
        End If
    Next i

    ' No matching entry found
    GetDailyAdmissionTotal = Empty
    Exit Function

ErrorHandler:
    ' On error, return Empty
    GetDailyAdmissionTotal = Empty
End Function

'===================================================================
' VALIDATE ADMISSION COUNT
'===================================================================
Public Function ValidateAdmissionCount(entryDate As Date, wardCode As String, _
    ByRef dailyTotal As Long, ByRef individualCount As Long, _
    ByRef errorMsg As String) As Boolean
    ' Compare daily total vs individual admission count
    '
    ' Args:
    '   entryDate: The date to validate
    '   wardCode: The ward code
    '   dailyTotal: (ByRef) Returns the daily total count
    '   individualCount: (ByRef) Returns the individual admission count
    '   errorMsg: (ByRef) Returns error message if validation fails
    '
    ' Returns:
    '   Boolean: True if counts match, False if mismatch or missing data

    On Error GoTo ErrorHandler

    errorMsg = ""

    ' Get daily total
    Dim dailyValue As Variant
    dailyValue = GetDailyAdmissionTotal(entryDate, wardCode)

    If IsEmpty(dailyValue) Then
        errorMsg = "No daily bed-state entry found for this date/ward"
        dailyTotal = 0
        individualCount = CountIndividualAdmissions(entryDate, wardCode)
        ValidateAdmissionCount = False
        Exit Function
    End If

    dailyTotal = CLng(dailyValue)
    individualCount = CountIndividualAdmissions(entryDate, wardCode)

    If dailyTotal <> individualCount Then
        errorMsg = "Mismatch: Daily total (" & dailyTotal & ") vs Individual count (" & individualCount & ")"
        ValidateAdmissionCount = False
    Else
        ValidateAdmissionCount = True
    End If

    Exit Function

ErrorHandler:
    errorMsg = "Error during validation: " & Err.Description
    ValidateAdmissionCount = False
End Function

'===================================================================
' GET MONTHLY VALIDATION REPORT
'===================================================================
Public Function GetMonthlyValidationReport(monthIndex As Long, wardCode As String, _
    reportYear As Long) As Variant
    ' Generate monthly validation report for ALL days in the month.
    '
    ' Args:
    '   monthIndex: Month number (1=January, 12=December)
    '   wardCode: The ward code to validate
    '   reportYear: The year (e.g., 2026)
    '
    ' Returns:
    '   Variant: 2D array with 5 columns per day:
    '     Col 1 - Date
    '     Col 2 - DailyTotal   (0 if NO ENTRY)
    '     Col 3 - IndividualCount (0 if NO ENTRY)
    '     Col 4 - Status: "OK" | "MISMATCH" | "NO ENTRY"
    '     Col 5 - Delta: DailyTotal - IndividualCount (0 if NO ENTRY)
    '   Returns Empty on error.

    On Error GoTo ErrorHandler

    ' Determine days in month
    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(reportYear, monthIndex + 1, 0))

    ' Allocate array for ALL days in the month (5 columns)
    Dim results() As Variant
    ReDim results(1 To daysInMonth, 1 To 5)

    Dim d As Long
    For d = 1 To daysInMonth
        Dim checkDate As Date
        checkDate = DateSerial(reportYear, monthIndex, d)

        results(d, 1) = checkDate

        Dim dailyValue As Variant
        dailyValue = GetDailyAdmissionTotal(checkDate, wardCode)

        If IsEmpty(dailyValue) Then
            ' No daily bed-state entry exists for this date
            results(d, 2) = 0           ' DailyTotal
            results(d, 3) = 0           ' IndividualCount
            results(d, 4) = "NO ENTRY"  ' Status
            results(d, 5) = 0           ' Delta
        Else
            Dim dailyTotal As Long
            Dim individualCount As Long
            dailyTotal = CLng(dailyValue)
            individualCount = CountIndividualAdmissions(checkDate, wardCode)

            results(d, 2) = dailyTotal
            results(d, 3) = individualCount
            results(d, 5) = dailyTotal - individualCount  ' Delta

            If dailyTotal = individualCount Then
                results(d, 4) = "OK"
            Else
                results(d, 4) = "MISMATCH"
            End If
        End If
    Next d

    GetMonthlyValidationReport = results
    Exit Function

ErrorHandler:
    GetMonthlyValidationReport = Empty
End Function

'===================================================================
' GET VALIDATION SUMMARY COUNTS
'===================================================================
Public Sub GetValidationSummary(results As Variant, _
    ByRef okCount As Long, ByRef mismatchCount As Long, _
    ByRef noEntryCount As Long)
    ' Counts OK / MISMATCH / NO ENTRY rows from a GetMonthlyValidationReport result.
    '
    ' Args:
    '   results: 2D array returned by GetMonthlyValidationReport
    '   okCount, mismatchCount, noEntryCount: (ByRef) output counts

    okCount = 0
    mismatchCount = 0
    noEntryCount = 0

    If IsEmpty(results) Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(results, 1)
        Select Case results(i, 4)
            Case "OK":       okCount = okCount + 1
            Case "MISMATCH": mismatchCount = mismatchCount + 1
            Case "NO ENTRY": noEntryCount = noEntryCount + 1
        End Select
    Next i
End Sub
