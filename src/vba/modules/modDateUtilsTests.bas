Attribute VB_Name = "modDateUtilsTests"
'==============================================================================
' Module: modDateUtilsTests
' Purpose: Unit tests for modDateUtils functions
'
' Usage: Run TestAll() from VBA Immediate Window or create a button
'
' Created: 2026-02-16
'==============================================================================

Option Explicit

'==============================================================================
' Main Test Runner
'==============================================================================
Public Sub TestAll()
    Dim passCount As Integer
    Dim failCount As Integer
    passCount = 0
    failCount = 0

    Debug.Print String(60, "=")
    Debug.Print "Running modDateUtils Test Suite"
    Debug.Print String(60, "=")
    Debug.Print ""

    ' Run all test suites
    TestParseDate passCount, failCount
    TestValidateDate passCount, failCount
    TestFormatDateDisplay passCount, failCount
    TestFormatDateStorage passCount, failCount
    TestGetDateFromString passCount, failCount

    ' Print summary
    Debug.Print ""
    Debug.Print String(60, "=")
    Debug.Print "Test Results: " & passCount & " passed, " & failCount & " failed"
    If failCount = 0 Then
        Debug.Print "SUCCESS - All tests passed!"
    Else
        Debug.Print "FAILURE - Some tests failed"
    End If
    Debug.Print String(60, "=")

    ' Show message box with results
    If failCount = 0 Then
        MsgBox "All tests passed! (" & passCount & " tests)", vbInformation, "Test Success"
    Else
        MsgBox failCount & " tests failed out of " & (passCount + failCount) & " total tests.", _
               vbExclamation, "Test Failure"
    End If
End Sub

'==============================================================================
' Test Suite: ParseDate
'==============================================================================
Private Sub TestParseDate(ByRef passCount As Integer, ByRef failCount As Integer)
    Debug.Print "Testing ParseDate..."

    Dim result As Variant
    Dim errMsg As String

    ' Test 1: Valid date
    result = modDateUtils.ParseDate("14/02/2026", errMsg)
    If IsDate(result) And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Valid date 14/02/2026"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Valid date 14/02/2026 - Expected date, got: " & result
    End If

    ' Test 2: Empty string
    result = modDateUtils.ParseDate("", errMsg)
    If IsEmpty(result) And errMsg <> "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Empty string returns error"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Empty string - Expected empty with error"
    End If

    ' Test 3: Invalid date format
    result = modDateUtils.ParseDate("99/99/9999", errMsg)
    If IsEmpty(result) And errMsg <> "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Invalid date 99/99/9999 returns error"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Invalid date - Expected empty with error"
    End If

    ' Test 4: Invalid day for month (Feb 30)
    result = modDateUtils.ParseDate("30/02/2026", errMsg)
    If IsEmpty(result) And InStr(errMsg, "Invalid day") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Feb 30 correctly rejected"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Feb 30 - Expected 'Invalid day' error"
    End If

    ' Test 5: Leap year valid (Feb 29, 2024)
    result = modDateUtils.ParseDate("29/02/2024", errMsg)
    If IsDate(result) And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Leap year Feb 29, 2024 accepted"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Leap year - Expected valid date"
    End If

    ' Test 6: Non-leap year invalid (Feb 29, 2025)
    result = modDateUtils.ParseDate("29/02/2025", errMsg)
    If IsEmpty(result) And InStr(errMsg, "Invalid day") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Non-leap year Feb 29, 2025 rejected"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Non-leap year - Expected 'Invalid day' error"
    End If

    ' Test 7: Year out of range (too old)
    result = modDateUtils.ParseDate("01/01/2019", errMsg)
    If IsEmpty(result) And InStr(errMsg, "2020") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Year 2019 rejected (before 2020)"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Year 2019 - Expected year range error"
    End If

    ' Test 8: Year out of range (too new)
    result = modDateUtils.ParseDate("01/01/2031", errMsg)
    If IsEmpty(result) And InStr(errMsg, "2030") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Year 2031 rejected (after 2030)"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Year 2031 - Expected year range error"
    End If

    ' Test 9: Non-numeric input
    result = modDateUtils.ParseDate("AA/BB/CCCC", errMsg)
    If IsEmpty(result) And errMsg <> "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Non-numeric input rejected"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Non-numeric - Expected error"
    End If

    ' Test 10: Valid edge case (Dec 31)
    result = modDateUtils.ParseDate("31/12/2026", errMsg)
    If IsDate(result) And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Dec 31, 2026 accepted"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Dec 31 - Expected valid date"
    End If

    Debug.Print ""
End Sub

'==============================================================================
' Test Suite: ValidateDate
'==============================================================================
Private Sub TestValidateDate(ByRef passCount As Integer, ByRef failCount As Integer)
    Debug.Print "Testing ValidateDate..."

    Dim result As Boolean
    Dim errMsg As String

    ' Test 1: Valid date in range
    result = modDateUtils.ValidateDate(DateSerial(2026, 2, 14), errMsg)
    If result = True And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Valid date 2026-02-14 accepted"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Valid date - Expected True"
    End If

    ' Test 2: Date before valid range
    result = modDateUtils.ValidateDate(DateSerial(2019, 12, 31), errMsg)
    If result = False And InStr(errMsg, "2020") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Date before 2020 rejected"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Date before range - Expected False"
    End If

    ' Test 3: Date after valid range
    result = modDateUtils.ValidateDate(DateSerial(2031, 1, 1), errMsg)
    If result = False And InStr(errMsg, "2030") > 0 Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Date after 2030 rejected"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Date after range - Expected False"
    End If

    ' Test 4: Boundary - First valid date (Jan 1, 2020)
    result = modDateUtils.ValidateDate(DateSerial(2020, 1, 1), errMsg)
    If result = True And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] First valid date Jan 1, 2020 accepted"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] First valid date - Expected True"
    End If

    ' Test 5: Boundary - Last valid date (Dec 31, 2030)
    result = modDateUtils.ValidateDate(DateSerial(2030, 12, 31), errMsg)
    If result = True And errMsg = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Last valid date Dec 31, 2030 accepted"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Last valid date - Expected True"
    End If

    Debug.Print ""
End Sub

'==============================================================================
' Test Suite: FormatDateDisplay
'==============================================================================
Private Sub TestFormatDateDisplay(ByRef passCount As Integer, ByRef failCount As Integer)
    Debug.Print "Testing FormatDateDisplay..."

    Dim result As String

    ' Test 1: Valid date formatting
    result = modDateUtils.FormatDateDisplay(DateSerial(2026, 2, 14))
    If result = "14/02/2026" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Date formats as dd/mm/yyyy"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Date format - Expected '14/02/2026', got: " & result
    End If

    ' Test 2: Invalid input (empty)
    result = modDateUtils.FormatDateDisplay(Empty)
    If result = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Empty input returns empty string"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Empty input - Expected empty string"
    End If

    ' Test 3: Single digit day and month
    result = modDateUtils.FormatDateDisplay(DateSerial(2026, 1, 5))
    If result = "05/01/2026" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Single digits formatted with leading zeros"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Single digits - Expected '05/01/2026', got: " & result
    End If

    Debug.Print ""
End Sub

'==============================================================================
' Test Suite: FormatDateStorage
'==============================================================================
Private Sub TestFormatDateStorage(ByRef passCount As Integer, ByRef failCount As Integer)
    Debug.Print "Testing FormatDateStorage..."

    Dim result As String

    ' Test 1: Valid date formatting for storage
    result = modDateUtils.FormatDateStorage(DateSerial(2026, 2, 14))
    If result = "2026-02-14" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Date formats as yyyy-mm-dd"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Storage format - Expected '2026-02-14', got: " & result
    End If

    ' Test 2: Invalid input (empty)
    result = modDateUtils.FormatDateStorage(Empty)
    If result = "" Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Empty input returns empty string"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Empty input - Expected empty string"
    End If

    Debug.Print ""
End Sub

'==============================================================================
' Test Suite: GetDateFromString
'==============================================================================
Private Sub TestGetDateFromString(ByRef passCount As Integer, ByRef failCount As Integer)
    Debug.Print "Testing GetDateFromString..."

    Dim result As Variant

    ' Test 1: Valid date string
    result = modDateUtils.GetDateFromString("14/02/2026")
    If IsDate(result) Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Valid date string returns date"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Valid date string - Expected date"
    End If

    ' Test 2: Invalid date string
    result = modDateUtils.GetDateFromString("99/99/9999")
    If IsEmpty(result) Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Invalid date string returns empty"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Invalid date string - Expected empty"
    End If

    ' Test 3: Date outside valid range
    result = modDateUtils.GetDateFromString("01/01/2019")
    If IsEmpty(result) Then
        passCount = passCount + 1
        Debug.Print "  [PASS] Date outside range returns empty"
    Else
        failCount = failCount + 1
        Debug.Print "  [FAIL] Date outside range - Expected empty"
    End If

    Debug.Print ""
End Sub
