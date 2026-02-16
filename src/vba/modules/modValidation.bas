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
    ' Generate monthly validation report for a specific ward
    '
    ' Args:
    '   monthIndex: Month number (1=January, 12=December)
    '   wardCode: The ward code to validate
    '   reportYear: The year (e.g., 2026)
    '
    ' Returns:
    '   Variant: 2D array with columns: Date, DailyTotal, IndividualCount, Status
    '           Returns Empty if no data found

    On Error GoTo ErrorHandler

    Dim results() As Variant
    Dim resultCount As Long
    resultCount = 0

    ' Determine days in month
    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(reportYear, monthIndex + 1, 0))

    ' Pre-allocate array (max size)
    ReDim results(1 To daysInMonth, 1 To 4)

    Dim d As Long
    For d = 1 To daysInMonth
        Dim checkDate As Date
        checkDate = DateSerial(reportYear, monthIndex, d)

        Dim dailyTotal As Long
        Dim individualCount As Long

        ' Get daily total
        Dim dailyValue As Variant
        dailyValue = GetDailyAdmissionTotal(checkDate, wardCode)

        ' Only include days with daily entries
        If Not IsEmpty(dailyValue) Then
            resultCount = resultCount + 1
            dailyTotal = CLng(dailyValue)
            individualCount = CountIndividualAdmissions(checkDate, wardCode)

            results(resultCount, 1) = checkDate
            results(resultCount, 2) = dailyTotal
            results(resultCount, 3) = individualCount

            If dailyTotal = individualCount Then
                results(resultCount, 4) = "OK"
            Else
                results(resultCount, 4) = "MISMATCH"
            End If
        End If
    Next d

    ' Resize to actual count
    If resultCount > 0 Then
        ReDim Preserve results(1 To resultCount, 1 To 4)
        GetMonthlyValidationReport = results
    Else
        GetMonthlyValidationReport = Empty
    End If

    Exit Function

ErrorHandler:
    ' On error, return Empty
    GetMonthlyValidationReport = Empty
End Function
