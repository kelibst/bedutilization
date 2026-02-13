"""
Phase 2: Inject VBA code into the workbook using win32com
Creates UserForms, standard modules, navigation buttons, and saves as .xlsm
"""
import os
import time
import json
from .config import WorkbookConfig

# ═══════════════════════════════════════════════════════════════════════════════
# VBA SOURCE CODE - STANDARD MODULES
# ═══════════════════════════════════════════════════════════════════════════════

VBA_MOD_CONFIG = '''
'###################################################################
'# MODULE: modConfig
'# PURPOSE: Hospital and ward configuration management
'###################################################################

Option Explicit

Public Const HOSPITAL_NAME As String = "HOHOE MUNICIPAL HOSPITAL"

'===================================================================
' WARD CONFIGURATION - reads from tblWardConfig table
'===================================================================
Public Function GetWardCount() As Long
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")
    GetWardCount = tbl.ListRows.Count
End Function

Public Function GetWardCodes() As Variant
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    Dim codes() As String
    ReDim codes(0 To tbl.ListRows.Count - 1)  ' 0-based for ComboBox compatibility

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        codes(i - 1) = tbl.ListRows(i).Range(1, 1).Value  ' WardCode column
    Next i

    GetWardCodes = codes
End Function

Public Function GetWardNames() As Variant
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    Dim names() As String
    ReDim names(0 To tbl.ListRows.Count - 1)  ' 0-based for ComboBox compatibility

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        names(i - 1) = tbl.ListRows(i).Range(1, 2).Value  ' WardName column
    Next i

    GetWardNames = names
End Function

Public Function GetWardByCode(wardCode As String) As Variant
    ' Returns array: (code, name, bedComplement, prevYearRemaining, isEmergency, displayOrder)
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range(1, 1).Value = wardCode Then
            Dim ward(1 To 6) As Variant
            ward(1) = tbl.ListRows(i).Range(1, 1).Value  ' Code
            ward(2) = tbl.ListRows(i).Range(1, 2).Value  ' Name
            ward(3) = tbl.ListRows(i).Range(1, 3).Value  ' BedComplement
            ward(4) = tbl.ListRows(i).Range(1, 4).Value  ' PrevYearRemaining
            ward(5) = tbl.ListRows(i).Range(1, 5).Value  ' IsEmergency
            ward(6) = tbl.ListRows(i).Range(1, 6).Value  ' DisplayOrder
            GetWardByCode = ward
            Exit Function
        End If
    Next i

    GetWardByCode = Null
End Function

Public Function GetReportYear() As Long
    GetReportYear = ThisWorkbook.Sheets("Control").Range("B5").Value
End Function

Public Sub ShowPreferencesInfo()
    ' Open preferences manager form
    frmPreferencesManager.Show
End Sub

Public Sub ExportPreferencesConfigButton()
    ExportPreferencesConfig
End Sub
'''

VBA_MOD_DATA_ACCESS = '''
'###################################################################
'# MODULE: modDataAccess
'# PURPOSE: Core data operations (save, retrieve, calculate)
'###################################################################

Option Explicit

'===================================================================
' COLUMN INDEX CONSTANTS
'===================================================================

' tblDaily columns
Public Const COL_DAILY_ENTRY_DATE As Integer = 1
Public Const COL_DAILY_MONTH As Integer = 2
Public Const COL_DAILY_WARD_CODE As Integer = 3
Public Const COL_DAILY_ADMISSIONS As Integer = 4
Public Const COL_DAILY_DISCHARGES As Integer = 5
Public Const COL_DAILY_DEATHS As Integer = 6
Public Const COL_DAILY_DEATHS_U24 As Integer = 7
Public Const COL_DAILY_TRANSFERS_IN As Integer = 8
Public Const COL_DAILY_TRANSFERS_OUT As Integer = 9
Public Const COL_DAILY_PREV_REMAINING As Integer = 10
Public Const COL_DAILY_REMAINING As Integer = 11
Public Const COL_DAILY_TIMESTAMP As Integer = 12

' tblAdmissions columns
Public Const COL_ADM_ID As Integer = 1
Public Const COL_ADM_DATE As Integer = 2
Public Const COL_ADM_MONTH As Integer = 3
Public Const COL_ADM_WARD_CODE As Integer = 4
Public Const COL_ADM_PATIENT_ID As Integer = 5
Public Const COL_ADM_PATIENT_NAME As Integer = 6
Public Const COL_ADM_AGE As Integer = 7
Public Const COL_ADM_AGE_UNIT As Integer = 8
Public Const COL_ADM_SEX As Integer = 9
Public Const COL_ADM_NHIS As Integer = 10
Public Const COL_ADM_TIMESTAMP As Integer = 11

' tblDeaths columns
Public Const COL_DEATH_ID As Integer = 1
Public Const COL_DEATH_DATE As Integer = 2
Public Const COL_DEATH_MONTH As Integer = 3
Public Const COL_DEATH_WARD_CODE As Integer = 4
Public Const COL_DEATH_FOLDER_NUM As Integer = 5
Public Const COL_DEATH_NAME As Integer = 6
Public Const COL_DEATH_AGE As Integer = 7
Public Const COL_DEATH_SEX As Integer = 8
Public Const COL_DEATH_NHIS As Integer = 9
Public Const COL_DEATH_CAUSE As Integer = 10
Public Const COL_DEATH_WITHIN_24HR As Integer = 11
Public Const COL_DEATH_AGE_UNIT As Integer = 12
Public Const COL_DEATH_TIMESTAMP As Integer = 13

'===================================================================
' REMAINING CALCULATION SYSTEM
'===================================================================
' The "Remaining" patient count is calculated as:
'   Remaining = PrevRemaining + Admissions - Discharges - Deaths - TransfersOut + TransfersIn
'
' CALCULATION FLOW:
' 1. SaveDailyEntry() writes data to tblDaily
' 2. CalculateRemainingForRow() calculates PrevRemaining and Remaining as VALUES (not formulas)
' 3. RecalculateSubsequentRows() cascades changes forward if middle rows are edited
'
' KEY FUNCTIONS:
' - GetLastRemainingForWard(wardCode, date): Finds previous remaining for context
' - CalculateRemainingForRow(rowIndex): Calculates one row's values
' - RecalculateAllRows(): Manual recalculation trigger (diagnostic)
'
' See CLAUDE.md "Remaining Calculation System" for full details.
'===================================================================

'===================================================================
' HELPER FUNCTIONS
'===================================================================

Private Function GetOrAddTableRow(tbl As ListObject) As ListRow
    ' Returns the first row if it's empty (seed row), otherwise adds a new row
    ' This avoids creating unnecessary empty rows in Excel tables

    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or _
           tbl.ListRows(1).Range(1, 1).Value = "" Then
            Set GetOrAddTableRow = tbl.ListRows(1)
            Exit Function
        End If
    End If

    Set GetOrAddTableRow = tbl.ListRows.Add
End Function

Private Function GenerateNextID(tbl As ListObject, prefix As String) As String
    ' Generates sequential IDs in format: [prefix]YYYY-#####
    ' prefix: Optional prefix (e.g., "D" for deaths, "" for admissions)
    ' Returns: Next available ID string

    Dim yr As Long
    yr = GetReportYear()

    ' Check if table is empty or has only seed row
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or _
           tbl.ListRows(1).Range(1, 1).Value = "" Then
            GenerateNextID = prefix & yr & "-00001"
            Exit Function
        End If
    End If

    ' Parse last ID and increment
    Dim lastID As String, dashPos As Long, num As Long
    lastID = CStr(tbl.ListRows(tbl.ListRows.Count).Range(1, 1).Value)
    dashPos = InStr(lastID, "-")

    If dashPos > 0 Then
        num = CLng(Mid(lastID, dashPos + 1)) + 1
    Else
        num = 1
    End If

    GenerateNextID = prefix & yr & "-" & Format(num, "00000")
End Function

'===================================================================
' CALCULATION OPERATIONS
'===================================================================

Public Function GetLastRemainingForWard(wardCode As String, beforeDate As Date) As Long
    '===================================================================
    ' CRITICAL: This function searches BACKWARD through all rows to find
    ' the most recent "Remaining" value for a ward before a given date.
    '
    ' BUG FIX (2026-02): Removed early exit condition in backward loop.
    ' Previously, the loop would exit as soon as it found ANY ward match,
    ' even if the date didn't match. This caused incorrect remaining values
    ' (e.g., Day 3 showing 25 instead of 23).
    '
    ' CORRECT BEHAVIOR: Must scan ALL rows backward until finding exact
    ' ward + date match, then return that row's Remaining value.
    '===================================================================

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    ' Check for empty table
    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
            GetLastRemainingForWard = GetPrevYearRemaining(wardCode)
            Exit Function
        End If
    End If

    ' Scan backward through ALL rows - no premature exits
    Dim foundRemaining As Long
    Dim foundDate As Date
    Dim hasMatch As Boolean

    foundDate = #1/1/1900#
    hasMatch = False

    Dim i As Long
    For i = tbl.ListRows.Count To 1 Step -1
        Dim rowWard As String
        Dim rowDate As Variant

        rowWard = tbl.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value
        rowDate = tbl.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value

        If rowWard = wardCode And IsDate(rowDate) Then
            Dim currentDate As Date
            currentDate = CDate(rowDate)
            If currentDate < beforeDate Then
                ' Found a matching prior entry - keep track of most recent
                If currentDate > foundDate Then
                    foundDate = currentDate
                    foundRemaining = CLng(tbl.ListRows(i).Range(1, COL_DAILY_REMAINING).Value)
                    hasMatch = True
                End If
            End If
        End If
    Next i

    ' Return the most recent match, or PrevYearRemaining if no match
    If hasMatch Then
        GetLastRemainingForWard = foundRemaining
    Else
        GetLastRemainingForWard = GetPrevYearRemaining(wardCode)
    End If
End Function

Public Function GetPrevYearRemaining(wardCode As String) As Long
    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets("Control")
    Dim tbl As ListObject
    Set tbl = wsCtrl.ListObjects("tblWardConfig")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range(1, 1).Value = wardCode Then
            GetPrevYearRemaining = CLng(tbl.ListRows(i).Range(1, 4).Value)
            Exit Function
        End If
    Next i
    GetPrevYearRemaining = 0
End Function

Public Function GetBedComplement(wardCode As String) As Long
    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets("Control")
    Dim tbl As ListObject
    Set tbl = wsCtrl.ListObjects("tblWardConfig")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range(1, 1).Value = wardCode Then
            GetBedComplement = CLng(tbl.ListRows(i).Range(1, 3).Value)
            Exit Function
        End If
    Next i
    GetBedComplement = 0
End Function

Public Function CheckDuplicateDaily(entryDate As Date, wardCode As String) As Long
    ' Returns the row index (within the table) if a duplicate exists, 0 otherwise
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim rowDate As Variant
        rowDate = tbl.ListRows(i).Range(1, COL_DAILY_ENTRY_DATE).Value
        If IsDate(rowDate) Then
            If CDate(rowDate) = entryDate And _
               tbl.ListRows(i).Range(1, COL_DAILY_WARD_CODE).Value = wardCode Then
                CheckDuplicateDaily = i
                Exit Function
            End If
        End If
    Next i
    CheckDuplicateDaily = 0
End Function

Public Sub CalculateRemainingForRow(targetRow As ListRow)
    ' UNIFIED CALCULATION FUNCTION - used by both form entry and manual edits
    ' This is the SINGLE source of truth for calculating PrevRemaining and Remaining
    On Error Resume Next

    ' Get values from the row
    Dim entryDate As Variant
    entryDate = targetRow.Range.Cells(1, COL_DAILY_ENTRY_DATE).Value
    If Not IsDate(entryDate) Then Exit Sub

    Dim wardCode As String
    wardCode = CStr(targetRow.Range.Cells(1, COL_DAILY_WARD_CODE).Value)
    If wardCode = "" Then Exit Sub

    ' Get previous remaining (looks backward in time)
    Dim prevRemaining As Long
    prevRemaining = GetLastRemainingForWard(wardCode, CDate(entryDate))

    ' Get input values from the row
    Dim admissions As Long, discharges As Long, deaths As Long
    Dim deathsU24 As Long, transIn As Long, transOut As Long

    admissions = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_ADMISSIONS).Value))
    discharges = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_DISCHARGES).Value))
    deaths = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_DEATHS).Value))
    deathsU24 = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_DEATHS_U24).Value))
    transIn = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_TRANSFERS_IN).Value))
    transOut = CLng(Val(targetRow.Range.Cells(1, COL_DAILY_TRANSFERS_OUT).Value))

    ' Calculate remaining
    ' Formula: Remaining = PrevRemaining + Admissions + TransfersIn - Discharges - Deaths - TransfersOut - DeathsUnder24Hrs
    Dim remaining As Long
    remaining = prevRemaining + admissions + transIn - discharges - deaths - transOut - deathsU24

    ' Update the calculated columns
    targetRow.Range.Cells(1, COL_DAILY_PREV_REMAINING).Value = prevRemaining
    targetRow.Range.Cells(1, COL_DAILY_REMAINING).Value = remaining
End Sub

Public Sub SaveDailyEntry(entryDate As Date, wardCode As String, _
    admissions As Long, discharges As Long, deaths As Long, _
    deathsU24 As Long, transIn As Long, transOut As Long)

    ' PERFORMANCE OPTIMIZATIONS
    Application.EnableEvents = False        ' Prevent event recursion
    Application.ScreenUpdating = False      ' 50-80% faster
    Application.Calculation = xlCalculationManual  ' No mid-save recalcs

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    ' Check if this is the seed row
    Dim useSeedRow As Boolean
    useSeedRow = False
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, COL_DAILY_ENTRY_DATE).Value) Then
            useSeedRow = True
        End If
    End If

    Dim existingRow As Long
    existingRow = CheckDuplicateDaily(entryDate, wardCode)

    Dim targetRow As ListRow
    If existingRow > 0 Then
        Set targetRow = tbl.ListRows(existingRow)
    ElseIf useSeedRow Then
        Set targetRow = tbl.ListRows(1)
    Else
        Set targetRow = tbl.ListRows.Add
    End If

    ' Write input values
    With targetRow.Range
        .Cells(1, COL_DAILY_ENTRY_DATE).Value = entryDate

        ' IMPORTANT: Date columns must be formatted explicitly
        ' Excel may store dates as text if format isn't set
        ' See initialize_date_formats() in phase2_vba.py for column setup
        .Cells(1, COL_DAILY_ENTRY_DATE).NumberFormat = "yyyy-mm-dd"

        .Cells(1, COL_DAILY_MONTH).Value = Month(entryDate)
        .Cells(1, COL_DAILY_WARD_CODE).Value = wardCode
        .Cells(1, COL_DAILY_ADMISSIONS).Value = admissions
        .Cells(1, COL_DAILY_DISCHARGES).Value = discharges
        .Cells(1, COL_DAILY_DEATHS).Value = deaths
        .Cells(1, COL_DAILY_DEATHS_U24).Value = deathsU24
        .Cells(1, COL_DAILY_TRANSFERS_IN).Value = transIn
        .Cells(1, COL_DAILY_TRANSFERS_OUT).Value = transOut
        .Cells(1, COL_DAILY_TIMESTAMP).Value = Now
        .Cells(1, COL_DAILY_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm"
    End With

    ' Calculate using unified function
    CalculateRemainingForRow targetRow

    ' Sort table
    SortDailyTable

    ' Find row index after sort (may have moved)
    Dim newRowIndex As Long
    For newRowIndex = 1 To tbl.ListRows.Count
        If tbl.ListRows(newRowIndex).Range(1, COL_DAILY_ENTRY_DATE).Value = entryDate And _
           tbl.ListRows(newRowIndex).Range(1, COL_DAILY_WARD_CODE).Value = wardCode Then
            Exit For
        End If
    Next newRowIndex

    ' Cascade recalculation
    RecalculateSubsequentRows tbl, newRowIndex, wardCode

    ' RESTORE Excel state
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate  ' Force ward sheets to update
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub RecalculateRow(tbl As ListObject, rowIndex As Long)
    ' Recalculate PrevRemaining and Remaining for a specific row
    ' This is called when user manually edits a row
    ' Simply delegates to the unified calculation function
    On Error Resume Next

    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Sub

    Dim targetRow As ListRow
    Set targetRow = tbl.ListRows(rowIndex)

    ' Use unified calculation function
    CalculateRemainingForRow targetRow
End Sub

Public Sub RecalculateSubsequentRows(tbl As ListObject, startRowIndex As Long, wardCode As String)
    ' OPTIMIZED: Early exit when different ward encountered
    On Error Resume Next

    If startRowIndex >= tbl.ListRows.Count Then Exit Sub

    Dim i As Long
    Dim processedCount As Long
    processedCount = 0

    For i = startRowIndex + 1 To tbl.ListRows.Count
        Dim rowWard As String
        rowWard = tbl.ListRows(i).Range.Cells(1, 3).Value

        If rowWard = wardCode Then
            ' Same ward - recalculate
            CalculateRemainingForRow tbl.ListRows(i)
            processedCount = processedCount + 1
        ElseIf processedCount > 0 Then
            ' Different ward & we've already processed some rows
            ' Table is sorted by ward, so we're done
            Exit For
        End If
    Next i
End Sub

Public Sub SortDailyTable()
    ' Sort tblDaily by WardCode, then EntryDate (ascending)
    ' This ensures MAXIFS finds the correct "previous" remaining value
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    ' Only sort if table has real data
    If tbl.ListRows.Count <= 1 Then Exit Sub
    If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Then Exit Sub

    ' Clear existing sort
    ws.Sort.SortFields.Clear

    ' Add sort fields: WardCode (col 3), then EntryDate (col 1)
    ws.Sort.SortFields.Add Key:=tbl.ListColumns("WardCode").DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=tbl.ListColumns("EntryDate").DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    ' Apply sort to table
    With ws.Sort
        .SetRange tbl.Range
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub RecalculateAllRows()
    ' Manual recalculation of all rows - useful for testing
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    ' Sort first
    SortDailyTable

    ' Recalculate each row
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) Then
            CalculateRemainingForRow tbl.ListRows(i)
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Recalculated " & tbl.ListRows.Count & " rows", vbInformation
End Sub

Public Sub VerifyCalculations()
    ' Diagnostic: Check if all calculations are correct
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    Dim errors As Long
    errors = 0

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim entryDate As Variant
        Dim wardCode As String

        entryDate = tbl.ListRows(i).Range(1, 1).Value
        wardCode = tbl.ListRows(i).Range(1, 3).Value

        If IsDate(entryDate) And wardCode <> "" Then
            ' Calculate expected
            Dim expectedPrev As Long
            expectedPrev = GetLastRemainingForWard(wardCode, CDate(entryDate))

            Dim actualPrev As Long
            actualPrev = CLng(tbl.ListRows(i).Range(1, 10).Value)

            If actualPrev <> expectedPrev Then
                errors = errors + 1
                Debug.Print "Row " & i & " ERROR: Ward=" & wardCode & _
                    " Date=" & Format(entryDate, "yyyy-mm-dd") & _
                    " Expected=" & expectedPrev & " Actual=" & actualPrev
            End If
        End If
    Next i

    If errors = 0 Then
        MsgBox "All " & tbl.ListRows.Count & " rows verified correct!", vbInformation
    Else
        MsgBox "Found " & errors & " errors. See Immediate Window.", vbExclamation
    End If
End Sub

Public Sub FixAllDateFormats()
    ' Fix date formatting issues across all data tables
    ' This procedure:
    ' 1. Applies proper date format to date columns
    ' 2. Converts text dates to proper date values
    ' 3. Reports what was fixed

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim fixCount As Long
    fixCount = 0
    Dim report As String
    report = "DATE FORMAT FIX REPORT" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf

    ' Fix DailyData table
    Dim wsDailyData As Worksheet
    Set wsDailyData = ThisWorkbook.Sheets("DailyData")
    Dim tblDaily As ListObject
    Set tblDaily = wsDailyData.ListObjects("tblDaily")

    report = report & "DailyData Table:" & vbCrLf
    ' EntryDate column (col 1)
    fixCount = fixCount + FixDateColumn(tblDaily, 1, "yyyy-mm-dd", "EntryDate")
    ' EntryTimestamp column (col 12)
    fixCount = fixCount + FixDateColumn(tblDaily, 12, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed " & fixCount & " date cells" & vbCrLf & vbCrLf

    ' Fix Admissions table
    Dim wsAdm As Worksheet
    Set wsAdm = ThisWorkbook.Sheets("Admissions")
    Dim tblAdm As ListObject
    Set tblAdm = wsAdm.ListObjects("tblAdmissions")

    Dim admFixCount As Long
    report = report & "Admissions Table:" & vbCrLf
    ' AdmissionDate column (col 2)
    admFixCount = FixDateColumn(tblAdm, 2, "yyyy-mm-dd", "AdmissionDate")
    ' EntryTimestamp column (col 11)
    admFixCount = admFixCount + FixDateColumn(tblAdm, 11, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed " & admFixCount & " date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + admFixCount

    ' Fix DeathsData table
    Dim wsDeaths As Worksheet
    Set wsDeaths = ThisWorkbook.Sheets("DeathsData")
    Dim tblDeaths As ListObject
    Set tblDeaths = wsDeaths.ListObjects("tblDeaths")

    Dim deathFixCount As Long
    report = report & "DeathsData Table:" & vbCrLf
    ' DateOfDeath column (col 2)
    deathFixCount = FixDateColumn(tblDeaths, 2, "yyyy-mm-dd", "DateOfDeath")
    ' EntryTimestamp column (col 13)
    deathFixCount = deathFixCount + FixDateColumn(tblDeaths, 13, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed " & deathFixCount & " date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + deathFixCount

    ' Fix TransfersData table
    Dim wsTrans As Worksheet
    Set wsTrans = ThisWorkbook.Sheets("TransfersData")
    Dim tblTrans As ListObject
    Set tblTrans = wsTrans.ListObjects("tblTransfers")

    Dim transFixCount As Long
    report = report & "TransfersData Table:" & vbCrLf
    ' TransferDate column (col 2)
    transFixCount = FixDateColumn(tblTrans, 2, "yyyy-mm-dd", "TransferDate")
    ' EntryTimestamp column (col 8)
    transFixCount = transFixCount + FixDateColumn(tblTrans, 8, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed " & transFixCount & " date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + transFixCount

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    report = report & String(50, "=") & vbCrLf
    report = report & "TOTAL: Fixed " & fixCount & " date cells across all tables"

    MsgBox report, vbInformation, "Date Format Fix Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error fixing date formats: " & Err.Description, vbCritical, "Error"
End Sub

Private Function FixDateColumn(tbl As ListObject, colIndex As Long, _
                               dateFormat As String, colName As String) As Long
    ' Fix a single date column in a table
    ' Returns: Number of cells fixed

    Dim fixedCount As Long
    fixedCount = 0

    ' First, apply format to entire column (including empty cells for future entries)
    On Error Resume Next
    tbl.ListColumns(colIndex).DataBodyRange.NumberFormat = dateFormat
    On Error GoTo 0

    ' Now fix any text dates or improperly formatted dates
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim cellVal As Variant
        cellVal = tbl.ListRows(i).Range(1, colIndex).Value

        ' Skip empty cells
        If Not IsEmpty(cellVal) And cellVal <> "" Then
            Dim originalVal As Variant
            originalVal = cellVal

            ' If it's stored as text but looks like a date, convert it
            If VarType(cellVal) = vbString Then
                If IsDate(cellVal) Then
                    ' Convert text to date
                    tbl.ListRows(i).Range(1, colIndex).Value = CDate(cellVal)
                    tbl.ListRows(i).Range(1, colIndex).NumberFormat = dateFormat
                    fixedCount = fixedCount + 1
                End If
            ElseIf IsDate(cellVal) Then
                ' It's already a date, but ensure format is correct
                ' Store as value then reformat to ensure consistency
                Dim tempDate As Date
                tempDate = CDate(cellVal)
                tbl.ListRows(i).Range(1, colIndex).Value = tempDate
                tbl.ListRows(i).Range(1, colIndex).NumberFormat = dateFormat

                ' Only count as fixed if format was different
                If tbl.ListRows(i).Range(1, colIndex).NumberFormat <> dateFormat Then
                    fixedCount = fixedCount + 1
                End If
            End If
        End If
    Next i

    FixDateColumn = fixedCount
End Function

Public Sub ImportFromOldWorkbook()
    ' Import data from a previous workbook (backward compatibility)
    ' Copies INPUT columns only and recalculates everything

    On Error GoTo ErrorHandler

    ' Prompt user to select old workbook
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.Title = "Select Previous Bed Utilization Workbook (.xlsm)"
    fd.Filters.Clear
    fd.Filters.Add "Excel Macro-Enabled Workbooks", "*.xlsm"
    fd.AllowMultiSelect = False

    If fd.Show <> -1 Then Exit Sub ' User cancelled

    Dim oldPath As String
    oldPath = fd.SelectedItems(1)

    ' Confirm with user
    If MsgBox("Import data from:" & vbCrLf & oldPath & vbCrLf & vbCrLf & _
              "This will copy all daily entries and recalculate. Continue?", _
              vbYesNo + vbQuestion, "Import Data") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Open old workbook
    Dim oldWB As Workbook
    Set oldWB = Workbooks.Open(oldPath, ReadOnly:=True, UpdateLinks:=False)

    ' Verify old workbook has DailyData sheet
    Dim oldWS As Worksheet
    On Error Resume Next
    Set oldWS = oldWB.Sheets("DailyData")
    On Error GoTo ErrorHandler

    If oldWS Is Nothing Then
        MsgBox "The selected workbook does not have a 'DailyData' sheet." & vbCrLf & _
               "Please select a valid Bed Utilization workbook.", vbExclamation
        oldWB.Close SaveChanges:=False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Verify old workbook has tblDaily table
    Dim oldTbl As ListObject
    On Error Resume Next
    Set oldTbl = oldWS.ListObjects("tblDaily")
    On Error GoTo ErrorHandler

    If oldTbl Is Nothing Then
        MsgBox "The DailyData sheet does not have a 'tblDaily' table." & vbCrLf & _
               "Please select a valid Bed Utilization workbook.", vbExclamation
        oldWB.Close SaveChanges:=False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Get new DailyData table
    Dim newWS As Worksheet
    Set newWS = ThisWorkbook.Sheets("DailyData")
    Dim newTbl As ListObject
    Set newTbl = newWS.ListObjects("tblDaily")

    Dim importCount As Long
    importCount = 0
    Dim useSeedRow As Boolean
    useSeedRow = False

    ' Check if new table has empty seed row
    If newTbl.ListRows.Count = 1 Then
        If IsEmpty(newTbl.ListRows(1).Range(1, 1).Value) Then
            useSeedRow = True
        End If
    End If

    ' ── NEW: Import Ward Configuration (PrevYearRemaining) ──
    ' This ensures that the starting values for calculations are preserved
    On Error Resume Next
    Dim oldConfigTbl As ListObject
    Set oldConfigTbl = oldWB.Sheets("Control").ListObjects("tblWardConfig")
    On Error GoTo ErrorHandler

    If Not oldConfigTbl Is Nothing Then
        Dim newConfigTbl As ListObject
        Set newConfigTbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")
        
        Dim r As Long
        For r = 1 To newConfigTbl.ListRows.Count
            Dim wCode As String
            wCode = newConfigTbl.ListRows(r).Range(1, 1).Value
            
            ' Find ward in old table
            Dim oldR As Long
            For oldR = 1 To oldConfigTbl.ListRows.Count
                If oldConfigTbl.ListRows(oldR).Range(1, 1).Value = wCode Then
                    ' Copy PrevYearRemaining (Col 4)
                    ' Only if old value is numeric
                    Dim oldVal As Variant
                    oldVal = oldConfigTbl.ListRows(oldR).Range(1, 4).Value
                    If IsNumeric(oldVal) Then
                        newConfigTbl.ListRows(r).Range(1, 4).Value = oldVal
                    End If
                    Exit For
                End If
            Next oldR
        Next r
    End If
    ' ────────────────────────────────────────────────────────

    ' Import each row from old workbook
    Dim i As Long
    For i = 1 To oldTbl.ListRows.Count
        ' Check if row has data
        Dim oldDate As Variant
        oldDate = oldTbl.ListRows(i).Range(1, 1).Value

        If IsDate(oldDate) Then
            ' Determine target row
            Dim newRow As ListRow
            If useSeedRow And importCount = 0 Then
                Set newRow = newTbl.ListRows(1)
            Else
                Set newRow = newTbl.ListRows.Add
            End If

            ' Copy INPUT columns only (1-9)
            ' Skip calculated columns (10-11) - will be recalculated
            With newRow.Range
                .Cells(1, 1).Value = oldTbl.ListRows(i).Range(1, 1).Value ' EntryDate
                .Cells(1, 2).Value = oldTbl.ListRows(i).Range(1, 2).Value ' Month
                .Cells(1, 3).Value = oldTbl.ListRows(i).Range(1, 3).Value ' WardCode
                .Cells(1, 4).Value = oldTbl.ListRows(i).Range(1, 4).Value ' Admissions
                .Cells(1, 5).Value = oldTbl.ListRows(i).Range(1, 5).Value ' Discharges
                .Cells(1, 6).Value = oldTbl.ListRows(i).Range(1, 6).Value ' Deaths
                .Cells(1, 7).Value = oldTbl.ListRows(i).Range(1, 7).Value ' DeathsU24
                .Cells(1, 8).Value = oldTbl.ListRows(i).Range(1, 8).Value ' TransfersIn
                .Cells(1, 9).Value = oldTbl.ListRows(i).Range(1, 9).Value ' TransfersOut
                ' Columns 10-11: Will be calculated after import
                .Cells(1, 12).Value = Now ' New timestamp
            End With

            importCount = importCount + 1
        End If
    Next i

    ' Close old workbook
    oldWB.Close SaveChanges:=False
    Set oldWB = Nothing

    ' Sort imported data
    SortDailyTable

    ' Recalculate all rows using unified function
    Dim j As Long
    For j = 1 To newTbl.ListRows.Count
        If Not IsEmpty(newTbl.ListRows(j).Range(1, 1).Value) Then
            CalculateRemainingForRow newTbl.ListRows(j)
        End If
    Next j

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Successfully imported " & importCount & " rows!" & vbCrLf & vbCrLf & _
           "All calculations have been recalculated automatically.", _
           vbInformation, "Import Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    If Not oldWB Is Nothing Then
        oldWB.Close SaveChanges:=False
    End If

    MsgBox "Error during import: " & Err.Description & vbCrLf & vbCrLf & _
           "Import cancelled.", vbExclamation, "Import Error"
End Sub

'===================================================================
' DATA SAVE OPERATIONS
'===================================================================

Public Sub SaveAdmission(admDate As Date, wardCode As String, _
    patientID As String, patientName As String, _
    age As Long, ageUnit As String, sex As String, nhis As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Admissions")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblAdmissions")

    ' Get row to use (seed row if empty, otherwise new row)
    Dim targetRow As ListRow
    Set targetRow = GetOrAddTableRow(tbl)

    ' Generate ID
    Dim newID As String
    newID = GenerateNextID(tbl, "")

    With targetRow.Range
        .Cells(1, COL_ADM_ID).Value = newID
        .Cells(1, COL_ADM_DATE).Value = admDate

        ' IMPORTANT: Date columns must be formatted explicitly
        ' Excel may store dates as text if format isn't set
        ' See initialize_date_formats() in phase2_vba.py for column setup
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
End Sub

Public Sub SaveDeath(deathDate As Date, wardCode As String, _
    folderNum As String, deceasedName As String, _
    age As Long, ageUnit As String, sex As String, _
    nhis As String, causeOfDeath As String, within24 As Boolean)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DeathsData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDeaths")

    ' Get row to use (seed row if empty, otherwise new row)
    Dim targetRow As ListRow
    Set targetRow = GetOrAddTableRow(tbl)

    ' Generate ID
    Dim newID As String
    newID = GenerateNextID(tbl, "D")

    With targetRow.Range
        .Cells(1, COL_DEATH_ID).Value = newID
        .Cells(1, COL_DEATH_DATE).Value = deathDate

        ' IMPORTANT: Date columns must be formatted explicitly
        ' Excel may store dates as text if format isn't set
        ' See initialize_date_formats() in phase2_vba.py for column setup
        .Cells(1, COL_DEATH_DATE).NumberFormat = "yyyy-mm-dd"

        .Cells(1, COL_DEATH_MONTH).Value = Month(deathDate)
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
End Sub
'''

VBA_MOD_REPORTS = '''
'###################################################################
'# MODULE: modReports
'# PURPOSE: Generate and refresh statistical reports
'###################################################################

Option Explicit

'===================================================================
' REPORT GENERATION FUNCTIONS
'===================================================================

Public Sub RefreshDeathsReport()
    Application.ScreenUpdating = False
    On Error GoTo cleanup

    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets("Deaths Report")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    ' Clear old data (keep headers)
    Dim monthStartRows(1 To 12) As Long
    Dim headerRows(1 To 12) As Long
    Dim r As Long
    r = 1
    Dim m As Long
    For m = 1 To 12
        monthStartRows(m) = r
        headerRows(m) = r + 1
        ' Clear data rows (rows 3 to 42 of each section = 40 rows)
        Dim dataStart As Long
        dataStart = r + 2
        wsReport.Range(wsReport.Cells(dataStart, 1), wsReport.Cells(dataStart + 39, 8)).ClearContents
        r = r + 42
    Next m

    ' Check if table has real data
    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            GoTo cleanup
        End If
    End If

    ' Populate from tblDeaths
    Dim sn(1 To 12) As Long
    For m = 1 To 12
        sn(m) = 0
    Next m

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim rowMonth As Variant
        rowMonth = tbl.ListRows(i).Range(1, 3).Value
        If Not IsEmpty(rowMonth) And IsNumeric(rowMonth) Then
            m = CLng(rowMonth)
            If m >= 1 And m <= 12 Then
                sn(m) = sn(m) + 1
                Dim writeRow As Long
                writeRow = monthStartRows(m) + 1 + sn(m)
                If sn(m) <= 40 Then
                    wsReport.Cells(writeRow, 1).Value = sn(m)
                    wsReport.Cells(writeRow, 2).Value = tbl.ListRows(i).Range(1, 5).Value  ' FolderNumber
                    wsReport.Cells(writeRow, 3).Value = tbl.ListRows(i).Range(1, 2).Value  ' DateOfDeath
                    wsReport.Cells(writeRow, 3).NumberFormat = "dd/mm/yyyy"
                    wsReport.Cells(writeRow, 4).Value = tbl.ListRows(i).Range(1, 6).Value  ' Name
                    wsReport.Cells(writeRow, 5).Value = tbl.ListRows(i).Range(1, 7).Value  ' Age
                    wsReport.Cells(writeRow, 6).Value = tbl.ListRows(i).Range(1, 9).Value  ' Sex
                    wsReport.Cells(writeRow, 7).Value = tbl.ListRows(i).Range(1, 4).Value  ' Ward
                    wsReport.Cells(writeRow, 8).Value = tbl.ListRows(i).Range(1, 10).Value ' NHIS
                End If
            End If
        End If
    Next i

cleanup:
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshCODSummary()
    Application.ScreenUpdating = False
    On Error GoTo cleanup2

    Dim wsCOD As Worksheet
    Set wsCOD = ThisWorkbook.Sheets("COD Summary")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    ' Clear old data (keep row 1-2 headers)
    If wsCOD.UsedRange.Rows.Count > 2 Then
        wsCOD.Range(wsCOD.Cells(3, 1), wsCOD.Cells(wsCOD.UsedRange.Rows.Count + 2, 14)).ClearContents
    End If

    ' Check if table has real data
    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            wsCOD.Cells(3, 1).Value = "(No death records found)"
            GoTo cleanup2
        End If
    End If

    ' Collect unique causes
    Dim causes As Object
    Set causes = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim cause As String
        cause = Trim(CStr(tbl.ListRows(i).Range(1, 11).Value))
        If cause <> "" And cause <> "0" Then
            If Not causes.Exists(cause) Then
                causes.Add cause, True
            End If
        End If
    Next i

    ' Write causes and COUNTIFS
    Dim causeKeys As Variant
    If causes.Count = 0 Then
        wsCOD.Cells(3, 1).Value = "(No causes recorded)"
        GoTo cleanup2
    End If

    causeKeys = causes.Keys
    Dim c As Long
    For c = 0 To UBound(causeKeys)
        Dim writeRow As Long
        writeRow = 3 + c
        wsCOD.Cells(writeRow, 1).Value = causeKeys(c)

        ' Count per month
        Dim m As Long
        For m = 1 To 12
            Dim cnt As Long
            cnt = 0
            For i = 1 To tbl.ListRows.Count
                If CLng(tbl.ListRows(i).Range(1, 3).Value) = m And _
                   Trim(CStr(tbl.ListRows(i).Range(1, 11).Value)) = causeKeys(c) Then
                    cnt = cnt + 1
                End If
            Next i
            wsCOD.Cells(writeRow, 1 + m).Value = cnt
        Next m

        ' Total
        wsCOD.Cells(writeRow, 14).Value = "=SUM(B" & writeRow & ":M" & writeRow & ")"
    Next c

cleanup2:
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshNonInsuredReport()
    Application.ScreenUpdating = False
    On Error GoTo cleanup3

    Dim wsRep As Worksheet
    Set wsRep = ThisWorkbook.Sheets("Non-Insured Report")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")

    ' Clear old data (keep row 1-2 headers)
    If wsRep.UsedRange.Rows.Count > 2 Then
        wsRep.Range(wsRep.Cells(3, 1), wsRep.Cells(wsRep.UsedRange.Rows.Count + 2, 10)).ClearContents
    End If

    ' Check if table has real data
    If tbl.ListRows.Count < 1 Then
        GoTo cleanup3
    End If
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            GoTo cleanup3
        End If
    End If

    Dim i As Long
    Dim r As Long
    r = 3
    Dim sn As Long
    sn = 0
    
    For i = 1 To tbl.ListRows.Count
        Dim nhis As String
        nhis = Trim(CStr(tbl.ListRows(i).Range(1, 10).Value))
        
        If UCase(nhis) = "NON-INSURED" Then
            sn = sn + 1
            wsRep.Cells(r, 1).Value = sn
            wsRep.Cells(r, 2).Value = tbl.ListRows(i).Range(1, 2).Value
            wsRep.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
            
            Dim dt As Variant
            dt = tbl.ListRows(i).Range(1, 2).Value
            If IsDate(dt) Then
                wsRep.Cells(r, 3).Value = Format(dt, "mmmm")
            End If
            
            wsRep.Cells(r, 4).Value = tbl.ListRows(i).Range(1, 3).Value
            wsRep.Cells(r, 5).Value = tbl.ListRows(i).Range(1, 4).Value
            wsRep.Cells(r, 6).Value = tbl.ListRows(i).Range(1, 6).Value
            wsRep.Cells(r, 7).Value = tbl.ListRows(i).Range(1, 7).Value & " " & tbl.ListRows(i).Range(1, 9).Value
            wsRep.Cells(r, 8).Value = tbl.ListRows(i).Range(1, 8).Value
            wsRep.Cells(r, 9).Value = "" 
            wsRep.Cells(r, 10).Value = nhis
            
            r = r + 1
        End If
    Next i
    
    If sn = 0 Then
        wsRep.Cells(3, 1).Value = "(No non-insured patients found)"
    End If

cleanup3:
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshAllReports()
    RefreshDeathsReport
    RefreshCODSummary
    RefreshNonInsuredReport
    MsgBox "All reports have been refreshed.", vbInformation, "Reports Updated"
End Sub
'''

VBA_MOD_NAVIGATION = '''
'###################################################################
'# MODULE: modNavigation
'# PURPOSE: Button event handlers for form navigation
'###################################################################

Option Explicit

'===================================================================
' FORM NAVIGATION HANDLERS
'===================================================================

Public Sub ShowDailyEntry()
    frmDailyEntry.Show
End Sub

Public Sub ShowAdmission()
    frmAdmission.Show
End Sub

Public Sub ShowDeath()
    frmDeath.Show
End Sub

Public Sub ShowAgesEntry()
    frmAgesEntry.Show
End Sub

Public Sub ShowRefreshReports()
    RefreshAllReports
End Sub

Public Sub ShowWardManager()
    frmWardManager.Show
End Sub

Public Sub ExportWardConfig()
    ExportWardsConfig
End Sub
'''

VBA_MOD_YEAREND = '''
'###################################################################
'# MODULE: modYearEnd
'# PURPOSE: Year-end operations, exports, and workbook rebuild
'###################################################################

Option Explicit

'===================================================================
' YEAR-END EXPORT FUNCTIONS
'===================================================================

Public Sub ExportCarryForward()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")
    Dim yr As Long
    yr = GetReportYear()

    ' Find Dec 31 remaining for each ward
    Dim wardCodes As Variant
    wardCodes = GetWardCodes()

    Dim jsonStr As String
    jsonStr = "{" & vbCrLf
    jsonStr = jsonStr & "  ""year"": " & yr & "," & vbCrLf
    jsonStr = jsonStr & "  ""wards"": {" & vbCrLf

    Dim w As Long
    For w = 0 To UBound(wardCodes)
        Dim wc As String
        wc = wardCodes(w)

        ' Find the last entry for this ward in December
        Dim lastRemaining As Long
        lastRemaining = 0
        Dim lastDate As Date
        lastDate = #1/1/1900#

        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            Dim rowDate As Variant
            rowDate = tbl.ListRows(i).Range(1, 1).Value
            If IsDate(rowDate) Then
                If tbl.ListRows(i).Range(1, 3).Value = wc Then
                    If CDate(rowDate) > lastDate Then
                        lastDate = CDate(rowDate)
                        lastRemaining = CLng(tbl.ListRows(i).Range(1, 11).Value)
                    End If
                End If
            End If
        Next i

        jsonStr = jsonStr & "    """ & wc & """: " & lastRemaining
        If w < UBound(wardCodes) Then jsonStr = jsonStr & ","
        jsonStr = jsonStr & vbCrLf
    Next w

    jsonStr = jsonStr & "  }" & vbCrLf
    jsonStr = jsonStr & "}"

    ' Save to file using helper function
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\carry_forward_" & yr & ".json"

    Call WriteJSONFile(filePath, jsonStr, "Carry-forward data exported to:")
End Sub

' Export current ward configuration to wards_config.json
Public Sub ExportWardsConfig()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    ' Build JSON structure
    Dim jsonStr As String
    jsonStr = "{" & vbCrLf
    jsonStr = jsonStr & "  ""_comment"": ""Ward Configuration for Bed Utilization System""," & vbCrLf
    jsonStr = jsonStr & "  ""_instructions"": [" & vbCrLf
    jsonStr = jsonStr & "    ""To add a new ward, copy an existing ward entry and modify the values.""," & vbCrLf
    jsonStr = jsonStr & "    ""After making changes, save this file and rebuild the workbook.""" & vbCrLf
    jsonStr = jsonStr & "  ]," & vbCrLf
    jsonStr = jsonStr & "  ""wards"": [" & vbCrLf

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        jsonStr = jsonStr & "    {" & vbCrLf
        jsonStr = jsonStr & "      ""code"": """ & tbl.ListRows(i).Range(1, 1).Value & """," & vbCrLf
        jsonStr = jsonStr & "      ""name"": """ & tbl.ListRows(i).Range(1, 2).Value & """," & vbCrLf
        jsonStr = jsonStr & "      ""bed_complement"": " & tbl.ListRows(i).Range(1, 3).Value & "," & vbCrLf
        jsonStr = jsonStr & "      ""prev_year_remaining"": " & tbl.ListRows(i).Range(1, 4).Value & "," & vbCrLf

        Dim isEmerg As Boolean
        isEmerg = tbl.ListRows(i).Range(1, 5).Value
        jsonStr = jsonStr & "      ""is_emergency"": " & LCase(CStr(isEmerg)) & "," & vbCrLf
        jsonStr = jsonStr & "      ""display_order"": " & tbl.ListRows(i).Range(1, 6).Value & vbCrLf
        jsonStr = jsonStr & "    }"
        If i < tbl.ListRows.Count Then jsonStr = jsonStr & ","
        jsonStr = jsonStr & vbCrLf
    Next i

    jsonStr = jsonStr & "  ]" & vbCrLf
    jsonStr = jsonStr & "}" & vbCrLf

    ' Save to wards_config.json
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\wards_config.json"

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonStr
    Close #fNum

    MsgBox "Ward configuration exported to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "To apply changes, rebuild the workbook:" & vbCrLf & _
           "python build_workbook.py --year " & GetReportYear(), _
           vbInformation, "Export Ward Configuration"
End Sub

Public Sub ExportPreferencesConfig()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    ' Read current values from table
    Dim showEmergencyRemaining As Boolean
    Dim subtractDeaths As Boolean

    ' Find rows by key (flexible ordering)
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))

        If key = "show_emergency_total_remaining" Then
            showEmergencyRemaining = CBool(tbl.ListRows(i).Range(1, 2).Value)
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            subtractDeaths = CBool(tbl.ListRows(i).Range(1, 2).Value)
        End If
    Next i

    ' Build JSON structure with proper formatting
    Dim jsonStr As String
    jsonStr = "{" & vbCrLf
    jsonStr = jsonStr & "  ""_comment"": ""Hospital Preferences for Bed Utilization System""," & vbCrLf
    jsonStr = jsonStr & "  ""_instructions"": [" & vbCrLf
    jsonStr = jsonStr & "    ""Configure hospital-specific behavioral preferences here.""," & vbCrLf
    jsonStr = jsonStr & "    ""After making changes, rebuild the workbook using: python build_workbook.py --year YYYY""," & vbCrLf
    jsonStr = jsonStr & "    ""Changing preferences mid-year may cause inconsistency between months.""" & vbCrLf
    jsonStr = jsonStr & "  ]," & vbCrLf
    jsonStr = jsonStr & "  ""version"": ""1.0""," & vbCrLf
    jsonStr = jsonStr & "  ""hospital_name"": ""HOHOE MUNICIPAL HOSPITAL""," & vbCrLf
    jsonStr = jsonStr & "  ""preferences"": {" & vbCrLf
    jsonStr = jsonStr & "    ""show_emergency_total_remaining"": " & LCase(CStr(showEmergencyRemaining)) & "," & vbCrLf
    jsonStr = jsonStr & "    ""subtract_deaths_under_24hrs_from_admissions"": " & LCase(CStr(subtractDeaths)) & vbCrLf
    jsonStr = jsonStr & "  }," & vbCrLf
    jsonStr = jsonStr & "  ""metadata"": {" & vbCrLf
    jsonStr = jsonStr & "    ""created_date"": """ & Format(Date, "yyyy-mm-dd") & """," & vbCrLf
    jsonStr = jsonStr & "    ""description"": ""Hospital-specific configuration preferences for bed utilization tracking""" & vbCrLf
    jsonStr = jsonStr & "  }" & vbCrLf
    jsonStr = jsonStr & "}" & vbCrLf

    ' Save to hospital_preferences.json
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\hospital_preferences.json"

    On Error GoTo ErrorHandler
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonStr
    Close #fNum

    ' Success message with rebuild instructions
    MsgBox "Hospital preferences exported to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "IMPORTANT: These changes will NOT take effect until you rebuild the workbook:" & vbCrLf & _
           "python build_workbook.py --year " & GetReportYear() & vbCrLf & vbCrLf & _
           "Preferences affect formulas and report structure, so rebuilding is required.", _
           vbInformation, "Export Hospital Preferences"
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting preferences: " & Err.Description, vbCritical, "Export Error"
End Sub

'===================================================================
' JSON HELPER FUNCTIONS
'===================================================================

Private Function WriteJSONFile(filePath As String, jsonContent As String, successMsg As String) As Boolean
    ' Centralized JSON file writing with error handling
    ' Returns True if successful, False otherwise
    On Error GoTo WriteError

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonContent
    Close #fNum

    If successMsg <> "" Then
        MsgBox successMsg & vbCrLf & filePath, vbInformation, "Export Successful"
    End If

    WriteJSONFile = True
    Exit Function

WriteError:
    WriteJSONFile = False
    MsgBox "Error writing JSON file: " & filePath & vbCrLf & Err.Description, vbCritical, "Export Error"
End Function

'===================================================================
' WORKBOOK REBUILD FUNCTIONS
'===================================================================
' Automated rebuild process with data preservation and validation

Private Function CheckRebuildPrerequisites() As String
    ' Returns empty string if all checks pass, otherwise returns error message
    ' Validates: Python installation, build scripts, config files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if build_workbook.py exists
    Dim buildScript As String
    buildScript = ThisWorkbook.Path & "\\build_workbook.py"
    If Not fso.FileExists(buildScript) Then
        CheckRebuildPrerequisites = "ERROR: build_workbook.py not found in workbook directory." & vbCrLf & _
                                    "Expected: " & buildScript
        Exit Function
    End If

    ' Check if Python is available
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\\python_check.txt"

    On Error Resume Next
    shell.Run "cmd /c python --version > """ & tempFile & """ 2>&1", 0, True
    On Error GoTo 0

    ' Read the result
    If fso.FileExists(tempFile) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(tempFile, 1)
        Dim pythonVersion As String
        If Not ts.AtEndOfStream Then
            pythonVersion = ts.ReadLine
        End If
        ts.Close
        fso.DeleteFile tempFile

        If InStr(LCase(pythonVersion), "python") = 0 Then
            CheckRebuildPrerequisites = "ERROR: Python not found or not in PATH." & vbCrLf & _
                                       "Please install Python and add it to your system PATH."
            Exit Function
        End If
    Else
        CheckRebuildPrerequisites = "ERROR: Unable to check Python installation."
        Exit Function
    End If

    ' Check if config files exist
    Dim wardsConfig As String
    wardsConfig = ThisWorkbook.Path & "\\config\\wards_config.json"
    If Not fso.FileExists(wardsConfig) Then
        CheckRebuildPrerequisites = "ERROR: config\\wards_config.json not found." & vbCrLf & _
                                    "Please export ward configuration first."
        Exit Function
    End If

    ' All checks passed
    CheckRebuildPrerequisites = ""
End Function

Public Sub RebuildWorkbookWithPreferences()
    ' Automated workbook rebuild with data preservation
    On Error GoTo ErrorHandler

    ' Step 1: Check prerequisites
    Dim prereqCheck As String
    prereqCheck = CheckRebuildPrerequisites()
    If prereqCheck <> "" Then
        MsgBox prereqCheck, vbCritical, "Rebuild Prerequisites Failed"
        Exit Sub
    End If

    ' Step 2: Confirm with user
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("This will rebuild the workbook with new preferences." & vbCrLf & vbCrLf & _
                         "Steps:" & vbCrLf & _
                         "1. Current workbook will be backed up" & vbCrLf & _
                         "2. Preferences will be exported to JSON" & vbCrLf & _
                         "3. Workbook will be rebuilt with Python" & vbCrLf & _
                         "4. You can import data from backup" & vbCrLf & vbCrLf & _
                         "Continue?", _
                         vbQuestion + vbYesNo, "Rebuild Workbook")
    If userResponse <> vbYes Then Exit Sub

    ' Step 3: Export preferences first
    Call ExportPreferencesConfig

    ' Step 4: Create backup of current workbook
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim currentPath As String
    Dim backupPath As String
    Dim timestamp As String
    currentPath = ThisWorkbook.FullName
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    backupPath = Replace(currentPath, ".xlsm", "_backup_" & timestamp & ".xlsm")

    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupPath
    Application.DisplayAlerts = True

    ' Step 5: Run Python rebuild
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim buildCmd As String
    Dim yearVal As String
    yearVal = CStr(GetReportYear())
    buildCmd = "cmd /c cd /d """ & ThisWorkbook.Path & """ && python build_workbook.py --year " & yearVal

    ' Create a log file for the build process
    Dim logFile As String
    logFile = ThisWorkbook.Path & "\\rebuild_log.txt"
    buildCmd = buildCmd & " > """ & logFile & """ 2>&1"

    ' Show progress message
    Application.StatusBar = "Rebuilding workbook... Please wait..."

    ' Run the command and wait
    Dim exitCode As Long
    exitCode = shell.Run(buildCmd, 0, True)

    Application.StatusBar = False

    ' Step 6: Check results
    If exitCode = 0 Then
        ' Success
        MsgBox "Rebuild successful!" & vbCrLf & vbCrLf & _
               "Backup saved to:" & vbCrLf & backupPath & vbCrLf & vbCrLf & _
               "NEXT STEPS:" & vbCrLf & _
               "1. This workbook will close" & vbCrLf & _
               "2. Open the new Bed_Utilization_" & yearVal & ".xlsm" & vbCrLf & _
               "3. Click 'IMPORT OLD WORKBOOK' button" & vbCrLf & _
               "4. Select the backup file to restore your data", _
               vbInformation, "Rebuild Complete"

        ' Close current workbook without saving (backup already created)
        ThisWorkbook.Close SaveChanges:=False
    Else
        ' Failed - read log file
        Dim errorMsg As String
        errorMsg = "Rebuild FAILED (Exit code: " & exitCode & ")" & vbCrLf & vbCrLf

        If fso.FileExists(logFile) Then
            Dim ts As Object
            Set ts = fso.OpenTextFile(logFile, 1)
            Dim logContent As String
            logContent = ts.ReadAll
            ts.Close

            ' Show last 500 characters of log
            If Len(logContent) > 500 Then
                logContent = "..." & Right(logContent, 500)
            End If
            errorMsg = errorMsg & "Build log:" & vbCrLf & logContent
        Else
            errorMsg = errorMsg & "No log file generated."
        End If

        MsgBox errorMsg, vbCritical, "Rebuild Failed"

        ' Keep current workbook open
        ' User can try again or manually fix issues
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    MsgBox "Error during rebuild: " & Err.Description & vbCrLf & vbCrLf & _
           "Your current workbook was not modified." & vbCrLf & _
           "Backup file: " & backupPath, _
           vbCritical, "Rebuild Error"
End Sub
'''

# ─── UserForm Code ───────────────────────────────────────────────────────────

VBA_FRM_DAILY_ENTRY_CODE = '''
Option Explicit

Private wardCodes As Variant
Private wardNames As Variant
Private isDirty As Boolean  ' Track if data has been modified
Private isLoading As Boolean  ' Prevent events during load

Private Sub UserForm_Initialize()
    isLoading = True
    isDirty = False

    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Populate month combo
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
    cmbMonth.ListIndex = Month(Date) - 1

    ' Set day spinner (1-31)
    spnDay.Min = 1
    spnDay.Max = 31
    spnDay.Value = Day(Date)
    txtDay.Value = CStr(Day(Date))

    ' Populate ward combo
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i

    ' Select first ward
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Initialize numeric fields to 0
    txtAdmissions.Value = "0"
    txtDischarges.Value = "0"
    txtDeaths.Value = "0"
    txtDeaths24.Value = "0"
    txtTransIn.Value = "0"
    txtTransOut.Value = "0"

    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
    UpdateRecentList
End Sub

Private Sub UpdateRecentList()
    lstRecent.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim i As Long
    For i = startRow To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
           tbl.ListRows(i).Range(1, 1).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 2).Value & " | " & _
                "Adm:" & tbl.ListRows(i).Range(1, 4).Value & _
                " Dis:" & tbl.ListRows(i).Range(1, 5).Value & _
                " Rem:" & tbl.ListRows(i).Range(1, 11).Value
        End If
    Next i
End Sub

Private Sub lstRecent_Click()
    If lstRecent.ListIndex < 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")

    ' Calculate actual row (last 10 entries)
    Dim startRow As Long
    startRow = tbl.ListRows.Count - 9
    If startRow < 1 Then startRow = 1

    Dim actualRow As Long
    actualRow = startRow + lstRecent.ListIndex

    If actualRow > tbl.ListRows.Count Then Exit Sub

    isLoading = True ' Prevent change events while loading
    On Error GoTo DateError

    ' Load the selected entry
    Dim entryDate As Date
    entryDate = CDate(tbl.ListRows(actualRow).Range(1, 1).Value)

    cmbMonth.ListIndex = Month(entryDate) - 1
    spnDay.Value = Day(entryDate)
    txtDay.Value = CStr(Day(entryDate))

    ' Load ward
    Dim wc As String
    wc = tbl.ListRows(actualRow).Range(1, 2).Value
    Dim j As Long
    For j = 0 To UBound(wardCodes)
        If wardCodes(j) = wc Then
            cmbWard.ListIndex = j
            Exit For
        End If
    Next j

    ' Load data fields
    txtAdmissions.Value = CStr(tbl.ListRows(actualRow).Range(1, 4).Value)
    txtDischarges.Value = CStr(tbl.ListRows(actualRow).Range(1, 5).Value)
    txtDeaths.Value = CStr(tbl.ListRows(actualRow).Range(1, 6).Value)
    txtDeaths24.Value = CStr(tbl.ListRows(actualRow).Range(1, 7).Value)
    txtTransIn.Value = CStr(tbl.ListRows(actualRow).Range(1, 8).Value)
    txtTransOut.Value = CStr(tbl.ListRows(actualRow).Range(1, 9).Value)

    isLoading = False
    UpdatePrevRemaining
    CalculateRemaining
    lblStatus.Caption = "Loaded entry for editing"
    lblStatus.ForeColor = &H0080FF ' Orange
    isDirty = False
    Exit Sub

DateError:
    isLoading = False
    MsgBox "Error loading entry: Invalid date format. Please contact support.", vbCritical, "Date Error"
    Exit Sub
End Sub

Private Sub cmbWard_Change()
    If isLoading Then Exit Sub

    ' Auto-save if data was modified before changing ward
    If isDirty Then
        If MsgBox("Save changes for current ward before switching?", vbYesNo + vbQuestion) = vbYes Then
            SaveCurrentEntry
        End If
        isDirty = False
    End If

    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub cmbMonth_Change()
    If isLoading Then Exit Sub
    UpdateDateFromControls
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub spnDay_Change()
    If isLoading Then Exit Sub
    txtDay.Value = CStr(spnDay.Value)
    UpdateDateFromControls
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub txtDay_Change()
    If isLoading Then Exit Sub
    On Error Resume Next
    Dim d As Long
    d = CLng(Val(txtDay.Value))
    If d >= 1 And d <= 31 Then
        spnDay.Value = d
    End If
End Sub

Private Sub UpdateDateFromControls()
    ' This just updates the internal date tracking
    ' The actual date is now derived from cmbMonth and spnDay
End Sub

Private Function GetSelectedDate() As Date
    On Error GoTo badDate
    Dim yr As Long, mo As Long, dy As Long
    yr = GetReportYear()
    mo = cmbMonth.ListIndex + 1
    dy = spnDay.Value

    ' Validate day for the month
    Dim maxDay As Long
    maxDay = Day(DateSerial(yr, mo + 1, 0))
    If dy > maxDay Then dy = maxDay

    GetSelectedDate = DateSerial(yr, mo, dy)
    Exit Function
badDate:
    GetSelectedDate = Date
End Function

Private Sub UpdatePrevRemaining()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim prevRem As Long
    prevRem = GetLastRemainingForWard(wc, entryDate)
    lblPrevRemaining.Caption = CStr(prevRem)

    Dim bc As Long
    bc = GetBedComplement(wc)
    lblBedComplement.Caption = CStr(bc)

    CalculateRemaining
End Sub

Private Sub CheckExistingEntry()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim existRow As Long
    existRow = CheckDuplicateDaily(entryDate, wc)

    isLoading = True  ' Prevent dirty flag while loading
    If existRow > 0 Then
        ' Load existing values
        Dim tbl As ListObject
        Set tbl = ThisWorkbook.Sheets("DailyData").ListObjects("tblDaily")
        txtAdmissions.Value = CStr(tbl.ListRows(existRow).Range(1, 4).Value)
        txtDischarges.Value = CStr(tbl.ListRows(existRow).Range(1, 5).Value)
        txtDeaths.Value = CStr(tbl.ListRows(existRow).Range(1, 6).Value)
        txtDeaths24.Value = CStr(tbl.ListRows(existRow).Range(1, 7).Value)
        txtTransIn.Value = CStr(tbl.ListRows(existRow).Range(1, 8).Value)
        txtTransOut.Value = CStr(tbl.ListRows(existRow).Range(1, 9).Value)
        lblStatus.Caption = "* Existing entry loaded"
        lblStatus.ForeColor = RGB(200, 100, 0)
    Else
        txtAdmissions.Value = "0"
        txtDischarges.Value = "0"
        txtDeaths.Value = "0"
        txtDeaths24.Value = "0"
        txtTransIn.Value = "0"
        txtTransOut.Value = "0"
            lblStatus.Caption = "New entry"
        lblStatus.ForeColor = RGB(100, 100, 100)
    End If
    isLoading = False
    isDirty = False
    CalculateRemaining
End Sub

Private Sub txtAdmissions_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDischarges_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDeaths_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtDeaths24_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTransIn_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub txtTransOut_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub

Private Sub CalculateRemaining()
    On Error Resume Next
    Dim prev As Long, adm As Long, dis As Long
    Dim dth As Long, dth24 As Long, ti As Long, tOut As Long, remVal As Long

    prev = CLng(Val(lblPrevRemaining.Caption))
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    dth24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    ' Formula: Remaining = Prev + Admissions + TransfersIn - (Discharges + Deaths + TransfersOut) - Deaths<24Hrs
    remVal = prev + adm + ti - dis - dth - tOut - dth24
    lblRemaining.Caption = CStr(remVal)

    If remVal < 0 Then
        lblRemaining.ForeColor = RGB(255, 0, 0)
    Else
        lblRemaining.ForeColor = RGB(0, 100, 0)
    End If
End Sub

Private Sub btnPrevDay_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    If spnDay.Value > 1 Then
        spnDay.Value = spnDay.Value - 1
    ElseIf cmbMonth.ListIndex > 0 Then
        cmbMonth.ListIndex = cmbMonth.ListIndex - 1
        spnDay.Value = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
    End If
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnNextDay_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    Dim maxDay As Long
    maxDay = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
    If spnDay.Value < maxDay Then
        spnDay.Value = spnDay.Value + 1
    ElseIf cmbMonth.ListIndex < 11 Then
        cmbMonth.ListIndex = cmbMonth.ListIndex + 1
        spnDay.Value = 1
    End If
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnToday_Click()
    ' Auto-save if dirty
    If isDirty Then SaveCurrentEntry

    isLoading = True
    cmbMonth.ListIndex = Month(Date) - 1
    spnDay.Value = Day(Date)
    txtDay.Value = CStr(spnDay.Value)
    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnSaveNext_Click()
    If SaveCurrentEntry() Then
        isDirty = False
        UpdateRecentList
        ' Move to next ward - preserve current ward index first
        Dim nextIdx As Long
        nextIdx = cmbWard.ListIndex + 1
        If nextIdx < cmbWard.ListCount Then
            isLoading = True  ' Prevent auto-save prompt
            cmbWard.ListIndex = nextIdx
            isLoading = False
            UpdatePrevRemaining
            CheckExistingEntry
        Else
            MsgBox "All wards completed for this date!", vbInformation
            ' Optionally advance to next day
            btnNextDay_Click
        End If
    End If
End Sub

Private Sub btnSaveNextDay_Click()
    If SaveCurrentEntry() Then
        isDirty = False
        UpdateRecentList
        ' Advance date to next day and reset to first ward
        isLoading = True
        Dim maxDay As Long
        maxDay = Day(DateSerial(GetReportYear(), cmbMonth.ListIndex + 2, 0))
        If spnDay.Value < maxDay Then
            spnDay.Value = spnDay.Value + 1
        ElseIf cmbMonth.ListIndex < 11 Then
            cmbMonth.ListIndex = cmbMonth.ListIndex + 1
            spnDay.Value = 1
        End If
        txtDay.Value = CStr(spnDay.Value)
        ' cmbWard.ListIndex = 0  <-- Removed to persist selection
        isLoading = False
        UpdatePrevRemaining
        CheckExistingEntry
    End If
End Sub

Private Sub btnSaveClose_Click()
    If SaveCurrentEntry() Then
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function SaveCurrentEntry() As Boolean
    SaveCurrentEntry = False

    ' Validate
    If cmbWard.ListIndex < 0 Then
        MsgBox "Please select a ward.", vbExclamation
        Exit Function
    End If

    Dim entryDate As Date
    entryDate = GetSelectedDate()

    Dim adm As Long, dis As Long, dth As Long
    Dim d24 As Long, ti As Long, tOut As Long
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    d24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    SaveDailyEntry entryDate, wc, adm, dis, dth, d24, ti, tOut

    lblStatus.Caption = "Saved: " & wardNames(cmbWard.ListIndex) & " - " & Format(entryDate, "dd/mm/yyyy")
    lblStatus.ForeColor = RGB(0, 128, 0)
    isDirty = False

    SaveCurrentEntry = True
End Function
'''

VBA_FRM_ADMISSION_CODE = '''
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
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
           tbl.ListRows(i).Range(1, 1).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 6).Value & " | " & _
                tbl.ListRows(i).Range(1, 4).Value & " | Age: " & _
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
    txtDate.Value = Format(tbl.ListRows(actualRow).Range(1, 1).Value, "dd/mm/yyyy")

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

    ' Load patient details
    txtPatientID.Value = tbl.ListRows(actualRow).Range(1, 2).Value
    txtPatientName.Value = tbl.ListRows(actualRow).Range(1, 4).Value
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

    Dim admDate As Date
    admDate = ParseDateAdm(txtDate.Value)
    If admDate = #1/1/1900# Then
        MsgBox "Please enter a valid date (dd/mm/yyyy).", vbExclamation
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

Private Sub UpdateAdmissionRow(rowIndex As Long, admDate As Date, wardCode As String, _
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
'''

VBA_FRM_AGES_ENTRY_CODE = '''
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

    ' Date defaults (same as Daily)
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
    cmbMonth.ListIndex = Month(Date) - 1

    spnDay.Min = 1
    spnDay.Max = 31
    spnDay.Value = Day(Date)
    txtDay.Value = CStr(Day(Date))

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
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
           tbl.ListRows(i).Range(1, 1).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
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
    dateVal = tbl.ListRows(actualRow).Range(1, 1).Value

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

    cmbMonth.ListIndex = Month(entryDate) - 1
    spnDay.Value = Day(entryDate)
    txtDay.Value = CStr(Day(entryDate))

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

    ' Read month and day from TEXT values (not control state)
    ' This prevents issues if controls lose state before save
    Dim monthName As String
    monthName = Trim(cmbMonth.Value)

    Dim dayText As String
    dayText = Trim(txtDay.Value)

    ' Convert month name to month number
    Dim monthNum As Long
    Select Case UCase(monthName)
        Case "JANUARY": monthNum = 1
        Case "FEBRUARY": monthNum = 2
        Case "MARCH": monthNum = 3
        Case "APRIL": monthNum = 4
        Case "MAY": monthNum = 5
        Case "JUNE": monthNum = 6
        Case "JULY": monthNum = 7
        Case "AUGUST": monthNum = 8
        Case "SEPTEMBER": monthNum = 9
        Case "OCTOBER": monthNum = 10
        Case "NOVEMBER": monthNum = 11
        Case "DECEMBER": monthNum = 12
        Case Else
            MsgBox "Invalid month selected: " & monthName & vbCrLf & "Please select a valid month.", vbExclamation, "Invalid Month"
            cmbMonth.SetFocus
            Exit Sub
    End Select

    ' Convert day text to number
    Dim dayNum As Long
    If Not IsNumeric(dayText) Then
        MsgBox "Invalid day value: " & dayText & vbCrLf & "Please enter a valid day (1-31).", vbExclamation, "Invalid Day"
        txtDay.SetFocus
        Exit Sub
    End If
    dayNum = CLng(dayText)

    If dayNum < 1 Or dayNum > 31 Then
        MsgBox "Day must be between 1 and 31. You entered: " & dayNum, vbExclamation, "Invalid Day"
        txtDay.SetFocus
        Exit Sub
    End If

    Dim yr As Long
    yr = GetReportYear()

    ' DIAGNOSTIC: Log values being used for date construction
    Debug.Print "=== Ages Entry Date Construction ==="
    Debug.Print "Year: " & yr
    Debug.Print "Month Name: " & monthName
    Debug.Print "Month Number: " & monthNum
    Debug.Print "Day Text: " & dayText
    Debug.Print "Day Number: " & dayNum
    Debug.Print "cmbMonth.ListIndex: " & cmbMonth.ListIndex
    Debug.Print "spnDay.Value: " & spnDay.Value

    Dim dt As Date
    On Error GoTo DateError
    dt = DateSerial(yr, monthNum, dayNum)
    On Error GoTo 0

    Debug.Print "Constructed Date: " & Format(dt, "yyyy-mm-dd")
    Debug.Print "===================================="

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

DateError:
    MsgBox "Error constructing date from selected month and day." & vbCrLf & _
           "Month: " & monthName & " (" & monthNum & ")" & vbCrLf & _
           "Day: " & dayNum & vbCrLf & _
           "Year: " & yr & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Date Error"
    Exit Sub
End Sub

Private Sub UpdateAgesRow(rowIndex As Long, admDate As Date, wardCode As String, _
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
        .Cells(1, COL_ADM_DATE).Value = admDate
        .Cells(1, COL_ADM_DATE).NumberFormat = "yyyy-mm-dd"
        .Cells(1, COL_ADM_MONTH).Value = Month(admDate)
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

Private Sub spnDay_Change()
    txtDay.Value = CStr(spnDay.Value)
End Sub

Private Sub txtDay_Change()
    If IsNumeric(txtDay.Value) Then
        Dim v As Long
        v = Val(txtDay.Value)
        If v >= 1 And v <= 31 Then spnDay.Value = v
    End If
End Sub
'''

VBA_FRM_DEATH_CODE = '''
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
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) And _
           tbl.ListRows(i).Range(1, 1).Value <> "" Then
            lstRecent.AddItem Format(tbl.ListRows(i).Range(1, 1).Value, "dd/mm/yyyy") & " | " & _
                tbl.ListRows(i).Range(1, 6).Value & " | " & _
                tbl.ListRows(i).Range(1, 3).Value & " | " & _
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
    txtDate.Value = Format(tbl.ListRows(actualRow).Range(1, 1).Value, "dd/mm/yyyy")

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

    ' Load patient details
    txtFolderNum.Value = tbl.ListRows(actualRow).Range(1, 2).Value
    txtName.Value = tbl.ListRows(actualRow).Range(1, 3).Value
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
    chkWithin24.Value = (tbl.ListRows(actualRow).Range(1, 5).Value = "Yes")

    ' Load cause
    cmbCause.Value = tbl.ListRows(actualRow).Range(1, 11).Value

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

    Dim deathDate As Date
    deathDate = ParseDateDth(txtDate.Value)
    If deathDate = #1/1/1900# Then
        MsgBox "Please enter a valid date (dd/mm/yyyy).", vbExclamation
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

Private Sub UpdateDeathRow(rowIndex As Long, deathDate As Date, wardCode As String, _
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
        .Cells(1, COL_DEATH_DATE).Value = deathDate
        .Cells(1, COL_DEATH_DATE).NumberFormat = "yyyy-mm-dd"
        .Cells(1, COL_DEATH_MONTH).Value = Month(deathDate)
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
'''

VBA_THISWORKBOOK = '''
Private Sub Workbook_Open()
    ' Navigate to Control sheet
    On Error Resume Next
    ThisWorkbook.Sheets("Control").Activate

    ' Hide data sheets
    ThisWorkbook.Sheets("DailyData").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Admissions").Visible = xlSheetHidden
    ThisWorkbook.Sheets("DeathsData").Visible = xlSheetHidden
    ThisWorkbook.Sheets("TransfersData").Visible = xlSheetHidden
End Sub
'''

VBA_SHEET_DAILYDATA = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Auto-recalculate when user manually edits input columns
    ' CRITICAL: Must be in DailyData worksheet module

    On Error GoTo ErrorHandler

    ' Check if events are enabled (prevents recursion)
    If Not Application.EnableEvents Then Exit Sub

    Dim tbl As ListObject
    Set tbl = Me.ListObjects("tblDaily")

    ' Check if change intersects table data body
    Dim intersectRange As Range
    Set intersectRange = Application.Intersect(Target, tbl.DataBodyRange)

    If Not intersectRange Is Nothing Then
        ' Disable events and screen updates for performance
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        ' Process each changed cell
        Dim cell As Range
        For Each cell In intersectRange
            ' Get column index within table (1-based)
            Dim colIdx As Long
            colIdx = cell.Column - tbl.Range.Column + 1

            ' Only trigger on input columns (4-9)
            If colIdx >= 4 And colIdx <= 9 Then
                ' Get row index within table
                Dim rowIdx As Long
                rowIdx = cell.Row - tbl.HeaderRowRange.Row

                If rowIdx >= 1 And rowIdx <= tbl.ListRows.Count Then
                    ' Recalculate this row
                    CalculateRemainingForRow tbl.ListRows(rowIdx)

                    ' Get ward code for cascading
                    Dim wardCode As String
                    wardCode = CStr(tbl.ListRows(rowIdx).Range.Cells(1, 3).Value)

                    ' Sort table (needed for PrevRemaining lookup)
                    SortDailyTable

                    ' Recalculate subsequent rows for this ward
                    RecalculateSubsequentRows tbl, rowIdx, wardCode

                    ' Exit after first change (avoid duplicate work)
                    Exit For
                End If
            End If
        Next cell

        ' Force ward sheets to recalculate
        Application.Calculate
    End If

ErrorHandler:
    ' Always restore Excel state even if error occurs
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        Debug.Print "Worksheet_Change Error: " & Err.Description
    End If
End Sub
'''


# ═══════════════════════════════════════════════════════════════════════════════
# USERFORM CREATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def create_daily_entry_form(vbproj):
    """Create the frmDailyEntry UserForm programmatically."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmDailyEntry"
    form.Properties("Caption").Value = "Daily Bed State Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 670

    d = form.Designer

    y = 12  # current Y position

    # Date selection: Month combo + Day spinner + navigation
    _add_label(d, "lblDateLabel", "Date:", 12, y, 40, 18)
    cmb_month = _add_combobox(d, "cmbMonth", 55, y, 100, 20)
    _add_label(d, "lblDayLabel", "Day:", 160, y, 30, 18)
    txt_day = _add_textbox(d, "txtDay", 195, y, 35, 20)
    # SpinButton for day
    spn = d.Controls.Add("Forms.SpinButton.1")
    spn.Name = "spnDay"
    spn.Left = 232
    spn.Top = y
    spn.Width = 18
    spn.Height = 20
    spn.Min = 1
    spn.Max = 31
    # Navigation buttons
    _add_button(d, "btnPrevDay", "< Prev", 260, y, 50, 20)
    _add_button(d, "btnNextDay", "Next >", 315, y, 50, 20)
    _add_button(d, "btnToday", "Today", 370, y, 42, 20)
    y += 28

    # Ward combo
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 120, 18)
    cmb = _add_combobox(d, "cmbWard", 140, y, 240, 22)
    y += 28

    # Previous Remaining
    _add_label(d, "lblPrevRemLabel", "Previous Remaining:", 12, y, 120, 18)
    lbl = _add_label(d, "lblPrevRemaining", "0", 140, y, 80, 18)
    lbl.Font.Bold = True
    lbl.Font.Size = 12
    lbl.ForeColor = 0x006400  # dark green

    _add_label(d, "lblBCLabel", "Bed Complement:", 230, y, 100, 18)
    lbl2 = _add_label(d, "lblBedComplement", "0", 340, y, 60, 18)
    lbl2.Font.Bold = True
    y += 32

    # Separator
    _add_label(d, "lblSep1", "", 12, y, 390, 1).BackColor = 0xC0C0C0
    y += 8

    # Numeric fields
    fields = [
        ("txtAdmissions", "Admissions:"),
        ("txtDischarges", "Discharges:"),
        ("txtDeaths", "Deaths:"),
        ("txtDeaths24", "Deaths < 24Hrs:"),
        ("txtTransIn", "Transfers In:"),
        ("txtTransOut", "Transfers Out:"),
    ]
    for name, caption in fields:
        _add_label(d, f"lbl{name}", caption, 12, y, 120, 18)
        _add_textbox(d, name, 140, y, 80, 20)
        y += 28

    # Separator
    _add_label(d, "lblSep2", "", 12, y, 390, 1).BackColor = 0xC0C0C0
    y += 8

    # Calculated Remaining
    _add_label(d, "lblRemLabel", "REMAINING:", 12, y, 120, 20)
    lbl3 = _add_label(d, "lblRemaining", "0", 140, y, 80, 20)
    lbl3.Font.Bold = True
    lbl3.Font.Size = 14
    lbl3.ForeColor = 0x006400
    y += 28

    # Status label
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons row 1
    _add_button(d, "btnSaveNext", "Save && Next Ward", 12, y, 130, 28)
    _add_button(d, "btnSaveNextDay", "Save && Next Day", 150, y, 120, 28)
    _add_button(d, "btnSaveClose", "Save && Close", 278, y, 100, 28)
    y += 32
    # Buttons row 2
    _add_button(d, "btnCancel", "Cancel", 12, y, 90, 28)
    y += 38

    # Recent entries list
    _add_label(d, "lblRecent", "Recent Daily Entries:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    # Inject code
    form.CodeModule.AddFromString(VBA_FRM_DAILY_ENTRY_CODE)


def create_admission_form(vbproj):
    """Create the frmAdmission UserForm."""
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAdmission"
    form.Properties("Caption").Value = "Patient Admission Record"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 520

    d = form.Designer
    y = 12

    # Date
    _add_label(d, "lblDateLabel", "Admission Date (dd/mm/yyyy):", 12, y, 160, 18)
    _add_textbox(d, "txtDate", 180, y, 120, 20)
    y += 28

    # Ward
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 180, y, 200, 22)
    y += 28

    # Patient ID
    _add_label(d, "lblPIDLabel", "Patient ID / Folder No:", 12, y, 160, 18)
    _add_textbox(d, "txtPatientID", 180, y, 200, 20)
    y += 28

    # Patient Name
    _add_label(d, "lblNameLabel", "Patient Name:", 12, y, 160, 18)
    _add_textbox(d, "txtPatientName", 180, y, 200, 20)
    y += 28

    # Age + Unit
    _add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    _add_textbox(d, "txtAge", 80, y, 60, 20)
    _add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex radio buttons
    _add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS radio buttons
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
    y += 28

    # Status
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    _add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    _add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    _add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Recent admissions list
    _add_label(d, "lblRecent", "Recent Admissions:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 380
    lst.Height = 100
    
    # Inject code
    form.CodeModule.AddFromString(VBA_FRM_ADMISSION_CODE)


def create_ages_entry_form(vbproj):
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmAgesEntry"
    form.Properties("Caption").Value = "Speed Ages Entry"
    form.Properties("Width").Value = 350
    form.Properties("Height").Value = 520

    d = form.Designer
    y = 12

    # Ward
    _add_label(d, "lblWard", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 100, y, 200, 22)
    y += 32

    # Date
    _add_label(d, "lblDate", "Date:", 12, y, 60, 18)
    _add_combobox(d, "cmbMonth", 80, y, 100, 22)
    _add_textbox(d, "txtDay", 190, y, 30, 20)
    spn = _add_spinner(d, "spnDay", 220, y, 15, 20)
    spn.Min = 1
    spn.Max = 31
    y += 32

    # Divider
    _add_label(d, "lblSep1", "", 12, y, 310, 1).BackColor = 0xC0C0C0
    y += 12

    # Age Entry Area
    _add_label(d, "lblAge", "AGE:", 12, y, 60, 20).Font.Bold = True
    _add_textbox(d, "txtAge", 80, y, 60, 24).Font.Size = 12

    _add_label(d, "lblUnit", "Unit:", 150, y+4, 40, 18)
    _add_combobox(d, "cmbAgeUnit", 195, y, 100, 22)
    y += 38

    # Sex
    _add_label(d, "lblSex", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 20, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 20, "grpSex")
    y += 28

    # Insurance
    _add_label(d, "lblIns", "Health Ins:", 12, y, 65, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 70, 20, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 160, y, 100, 20, "grpNHIS")
    y += 32

    # Status
    lbl = _add_label(d, "lblStatus", "Ready", 12, y, 310, 20)
    lbl.Font.Bold = True
    lbl.ForeColor = 0x808080 # Gray
    y += 24

    # Buttons
    btnSave = _add_button(d, "btnSave", "Save Entry (Enter)", 12, y, 140, 30)
    btnSave.Default = True
    _add_button(d, "btnClose", "Close", 160, y, 100, 30)
    y += 38

    # Recent entries list
    _add_label(d, "lblRecent", "Recent Age Entries:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 310
    lst.Height = 100

    # Inject
    form.CodeModule.AddFromString(VBA_FRM_AGES_ENTRY_CODE)


def create_death_form(vbproj):
    """Create the frmDeath UserForm."""
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmDeath"
    form.Properties("Caption").Value = "Death Record Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 620

    d = form.Designer
    y = 12

    # Date
    _add_label(d, "lblDateLabel", "Date of Death (dd/mm/yyyy):", 12, y, 170, 18)
    _add_textbox(d, "txtDate", 190, y, 120, 20)
    y += 28

    # Ward
    _add_label(d, "lblWardLabel", "Ward:", 12, y, 60, 18)
    _add_combobox(d, "cmbWard", 190, y, 200, 22)
    y += 28

    # Folder Number
    _add_label(d, "lblFolderLabel", "Folder Number:", 12, y, 170, 18)
    _add_textbox(d, "txtFolderNum", 190, y, 200, 20)
    y += 28

    # Name
    _add_label(d, "lblNameLabel", "Name of Deceased:", 12, y, 170, 18)
    _add_textbox(d, "txtName", 190, y, 200, 20)
    y += 28

    # Age + Unit
    _add_label(d, "lblAgeLabel", "Age:", 12, y, 60, 18)
    _add_textbox(d, "txtAge", 80, y, 60, 20)
    _add_combobox(d, "cmbAgeUnit", 150, y, 90, 22)
    y += 32

    # Sex
    _add_label(d, "lblSexLabel", "Sex:", 12, y, 60, 18)
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18, "grpSex")
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18, "grpSex")
    y += 28

    # NHIS
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18, "grpNHIS")
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18, "grpNHIS")
    y += 28

    # Death within 24hrs checkbox
    chk = d.Controls.Add("Forms.CheckBox.1")
    chk.Name = "chkWithin24"
    chk.Caption = "Death within 24 hours of admission"
    chk.Left = 12
    chk.Top = y
    chk.Width = 250
    chk.Height = 18
    y += 28

    # Cause of Death
    _add_label(d, "lblCauseLabel", "Cause of Death:", 12, y, 170, 18)
    cmb = d.Controls.Add("Forms.ComboBox.1")
    cmb.Name = "cmbCause"
    cmb.Left = 190
    cmb.Top = y
    cmb.Width = 200
    cmb.Height = 22
    cmb.Style = 0  # fmStyleDropDownCombo (allows free text)
    y += 32

    # Status
    _add_label(d, "lblStatus", "", 12, y, 390, 18)
    y += 24

    # Buttons
    _add_button(d, "btnSaveNew", "Save && New", 12, y, 110, 30)
    _add_button(d, "btnSaveClose", "Save && Close", 130, y, 110, 30)
    _add_button(d, "btnCancel", "Cancel", 250, y, 90, 30)
    y += 38

    # Recent deaths list
    _add_label(d, "lblRecent", "Recent Deaths:", 12, y, 150, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    form.CodeModule.AddFromString(VBA_FRM_DEATH_CODE)


# ═══════════════════════════════════════════════════════════════════════════════
# PREFERENCES MANAGER FORM
# ═══════════════════════════════════════════════════════════════════════════════

VBA_FRM_PREFERENCES_MANAGER_CODE = '''
Option Explicit

Private Sub UserForm_Initialize()
    LoadPreferences
End Sub

Private Sub LoadPreferences()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    ' Load checkbox values by matching keys
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))

        If key = "show_emergency_total_remaining" Then
            chkShowEmergencyRemaining.Value = CBool(tbl.ListRows(i).Range(1, 2).Value)
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            chkSubtractDeaths.Value = CBool(tbl.ListRows(i).Range(1, 2).Value)
        End If
    Next i
End Sub

Private Sub btnSave_Click()
    ' Save changes to table
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    ' Update table values
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))

        If key = "show_emergency_total_remaining" Then
            tbl.ListRows(i).Range(1, 2).Value = chkShowEmergencyRemaining.Value
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            tbl.ListRows(i).Range(1, 2).Value = chkSubtractDeaths.Value
        End If
    Next i

    MsgBox "Preferences saved to table!" & vbCrLf & vbCrLf & _
           "Next step: Click 'Export to JSON' button to save to file, " & _
           "then rebuild the workbook for changes to take effect.", _
           vbInformation, "Preferences Saved"

    Unload Me
End Sub

Private Sub btnExport_Click()
    ExportPreferencesConfig
End Sub

Private Sub btnSaveRebuild_Click()
    ' Save to table, then automatically rebuild workbook
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    ' Update table values
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))

        If key = "show_emergency_total_remaining" Then
            tbl.ListRows(i).Range(1, 2).Value = chkShowEmergencyRemaining.Value
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            tbl.ListRows(i).Range(1, 2).Value = chkSubtractDeaths.Value
        End If
    Next i

    ' Close form
    Unload Me

    ' Trigger automated rebuild
    RebuildWorkbookWithPreferences
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
'''

# ═══════════════════════════════════════════════════════════════════════════════
# WARD MANAGER FORM
# ═══════════════════════════════════════════════════════════════════════════════

VBA_FRM_WARD_MANAGER_CODE = '''
Option Explicit

Private Sub UserForm_Initialize()
    LoadWards
End Sub

Private Sub LoadWards()
    lstWards.Clear
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim dispText As String
        dispText = tbl.ListRows(i).Range(1, 1).Value & " - " & _
                   tbl.ListRows(i).Range(1, 2).Value & " (" & _
                   tbl.ListRows(i).Range(1, 3).Value & " beds)"
        lstWards.AddItem dispText
    Next i

    If lstWards.ListCount > 0 Then lstWards.ListIndex = 0
End Sub

Private Sub lstWards_Click()
    If lstWards.ListIndex < 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")
    Dim idx As Long
    idx = lstWards.ListIndex + 1

    txtCode.Value = tbl.ListRows(idx).Range(1, 1).Value
    txtName.Value = tbl.ListRows(idx).Range(1, 2).Value
    txtBeds.Value = tbl.ListRows(idx).Range(1, 3).Value
    txtPrevRemaining.Value = tbl.ListRows(idx).Range(1, 4).Value
    chkEmergency.Value = tbl.ListRows(idx).Range(1, 5).Value
    txtDisplayOrder.Value = tbl.ListRows(idx).Range(1, 6).Value
End Sub

Private Sub btnNew_Click()
    txtCode.Value = ""
    txtName.Value = ""
    txtBeds.Value = "0"
    txtPrevRemaining.Value = "0"
    chkEmergency.Value = False
    txtDisplayOrder.Value = lstWards.ListCount + 1
    txtCode.SetFocus
    lstWards.ListIndex = -1
End Sub

Private Sub btnSave_Click()
    ' Validate inputs
    If Trim(txtCode.Value) = "" Then
        MsgBox "Please enter a ward code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If

    If Trim(txtName.Value) = "" Then
        MsgBox "Please enter a ward name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtBeds.Value) Or CLng(txtBeds.Value) < 0 Then
        MsgBox "Bed complement must be a non-negative number.", vbExclamation
        txtBeds.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtDisplayOrder.Value) Or CLng(txtDisplayOrder.Value) < 1 Then
        MsgBox "Display order must be a positive number.", vbExclamation
        txtDisplayOrder.SetFocus
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    Dim wardRow As Long
    wardRow = lstWards.ListIndex + 1

    ' Check if this is a new ward (no selection) or update
    If lstWards.ListIndex < 0 Then
        ' New ward - check for duplicate code
        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            If UCase(Trim(tbl.ListRows(i).Range(1, 1).Value)) = UCase(Trim(txtCode.Value)) Then
                MsgBox "Ward code '" & txtCode.Value & "' already exists.", vbExclamation
                Exit Sub
            End If
        Next i

        ' Add new row
        tbl.ListRows.Add
        wardRow = tbl.ListRows.Count
    End If

    ' Save values
    tbl.ListRows(wardRow).Range(1, 1).Value = Trim(txtCode.Value)
    tbl.ListRows(wardRow).Range(1, 2).Value = Trim(txtName.Value)
    tbl.ListRows(wardRow).Range(1, 3).Value = CLng(txtBeds.Value)
    tbl.ListRows(wardRow).Range(1, 4).Value = CLng(txtPrevRemaining.Value)
    tbl.ListRows(wardRow).Range(1, 5).Value = chkEmergency.Value
    tbl.ListRows(wardRow).Range(1, 6).Value = CLng(txtDisplayOrder.Value)

    MsgBox "Ward saved successfully!" & vbCrLf & vbCrLf & _
           "Don't forget to export the configuration and rebuild the workbook.", _
           vbInformation

    LoadWards
End Sub

Private Sub btnDelete_Click()
    If lstWards.ListIndex < 0 Then
        MsgBox "Please select a ward to delete.", vbExclamation
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete this ward?" & vbCrLf & _
                      "This will NOT delete the ward sheet or data." & vbCrLf & vbCrLf & _
                      "You must rebuild the workbook to remove the ward sheet.", _
                      vbQuestion + vbYesNo, "Confirm Delete")

    If response = vbNo Then Exit Sub

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    tbl.ListRows(lstWards.ListIndex + 1).Delete

    LoadWards
    btnNew_Click
End Sub

Private Sub btnExport_Click()
    ExportWardsConfig
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''

def create_ward_manager_form(vbproj):
    """Create the frmWardManager UserForm."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmWardManager"
    form.Properties("Caption").Value = "Manage Ward Configuration"
    form.Properties("Width").Value = 520
    form.Properties("Height").Value = 400

    d = form.Designer

    y = 12

    # Title
    lbl = _add_label(d, "lblTitle", "Ward Configuration Manager", 12, y, 490, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = _add_label(d, "lblInstructions",
                     "Add or edit wards below. Click 'Export Config' to save to JSON, then rebuild the workbook.",
                     12, y, 490, 30)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 38

    # Ward list (left side)
    _add_label(d, "lblWards", "Wards:", 12, y, 180, 18)
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstWards"
    lst.Left = 12
    lst.Top = y + 22
    lst.Width = 180
    lst.Height = 240

    # Ward details (right side)
    x2 = 205
    _add_label(d, "lblDetails", "Ward Details:", x2, y, 200, 18)
    y2 = y + 22

    # Code
    _add_label(d, "lblCode", "Code:", x2, y2, 80, 18)
    _add_textbox(d, "txtCode", x2 + 105, y2, 100, 20)
    y2 += 26

    # Name
    _add_label(d, "lblName", "Name:", x2, y2, 80, 18)
    _add_textbox(d, "txtName", x2 + 105, y2, 200, 20)
    y2 += 26

    # Bed Complement
    _add_label(d, "lblBeds", "Bed Complement:", x2, y2, 100, 18)
    _add_textbox(d, "txtBeds", x2 + 105, y2, 60, 20)
    y2 += 26

    # Previous Year Remaining
    _add_label(d, "lblPrevRem", "Prev Year Remaining:", x2, y2, 100, 18)
    _add_textbox(d, "txtPrevRemaining", x2 + 105, y2, 60, 20)
    y2 += 26

    # Emergency checkbox
    chk = d.Controls.Add("Forms.CheckBox.1")
    chk.Name = "chkEmergency"
    chk.Caption = "Emergency Ward"
    chk.Left = x2
    chk.Top = y2
    chk.Width = 120
    chk.Height = 18
    y2 += 26

    # Display Order
    _add_label(d, "lblOrder", "Display Order:", x2, y2, 100, 18)
    _add_textbox(d, "txtDisplayOrder", x2 + 105, y2, 60, 20)
    y2 += 32

    # Buttons (right side)
    _add_button(d, "btnNew", "New Ward", x2, y2, 90, 28)
    _add_button(d, "btnSave", "Save", x2 + 95, y2, 90, 28)
    _add_button(d, "btnDelete", "Delete", x2 + 190, y2, 80, 28)

    # Bottom buttons
    y_bottom = 350
    _add_button(d, "btnExport", "Export Config to JSON", 12, y_bottom, 150, 28)
    _add_button(d, "btnClose", "Close", 390, y_bottom, 110, 28)

    form.CodeModule.AddFromString(VBA_FRM_WARD_MANAGER_CODE)


def create_preferences_manager_form(vbproj):
    """Create the frmPreferencesManager UserForm."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmPreferencesManager"
    form.Properties("Caption").Value = "Hospital Preferences Configuration"
    form.Properties("Width").Value = 500
    form.Properties("Height").Value = 370

    d = form.Designer
    y = 12

    # Title
    lbl = _add_label(d, "lblTitle", "Hospital Preferences", 12, y, 470, 20)
    lbl.Font.Bold = True
    lbl.Font.Size = 14
    lbl.TextAlign = 2  # center
    y += 28

    # Instructions
    lbl = _add_label(d, "lblInstructions",
                     "Configure hospital-specific preferences. After saving and exporting, rebuild the workbook for changes to take effect.",
                     12, y, 470, 35)
    lbl.ForeColor = 0x808080
    lbl.WordWrap = True
    y += 45

    # Warning frame
    lbl = _add_label(d, "lblWarning",
                     "WARNING: These preferences affect formulas and report structure. Changes require workbook rebuild!",
                     12, y, 470, 30)
    lbl.ForeColor = 0x0000C0  # Dark red
    lbl.WordWrap = True
    lbl.Font.Bold = True
    y += 40

    # Preference checkboxes
    chk1 = d.Controls.Add("Forms.CheckBox.1")
    chk1.Name = "chkShowEmergencyRemaining"
    chk1.Caption = "Show 'Emergency Total Remaining' row in Monthly Summary"
    chk1.Left = 20
    chk1.Top = y
    chk1.Width = 450
    chk1.Height = 18
    y += 30

    chk2 = d.Controls.Add("Forms.CheckBox.1")
    chk2.Name = "chkSubtractDeaths"
    chk2.Caption = "Subtract deaths under 24hrs from monthly admission totals"
    chk2.Left = 20
    chk2.Top = y
    chk2.Width = 450
    chk2.Height = 18
    y += 50

    # Buttons (arranged in 2 rows)
    y_buttons = 240
    # Top row - Primary actions
    _add_button(d, "btnSave", "Save to Table", 20, y_buttons, 140, 32)
    _add_button(d, "btnSaveRebuild", "Save & Rebuild", 170, y_buttons, 140, 32)
    _add_button(d, "btnCancel", "Cancel", 360, y_buttons, 120, 32)
    # Bottom row - Export only
    y_buttons += 42
    _add_button(d, "btnExport", "Export to JSON (without rebuild)", 20, y_buttons, 310, 28)

    form.CodeModule.AddFromString(VBA_FRM_PREFERENCES_MANAGER_CODE)


# ─── Helper functions for control creation ───────────────────────────────────

def _add_label(designer, name, caption, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.Label.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_textbox(designer, name, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.TextBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_combobox(designer, name, left, top, width, height, style=0):
    ctrl = designer.Controls.Add("Forms.ComboBox.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Style = style  # 0=DropDownCombo (default), 2=DropDownList
    return ctrl


def _add_optionbutton(designer, name, caption, left, top, width, height, group_name=None):
    ctrl = designer.Controls.Add("Forms.OptionButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    if group_name:
        ctrl.GroupName = group_name
    return ctrl


def _add_spinner(designer, name, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.SpinButton.1")
    ctrl.Name = name
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_button(designer, name, caption, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.CommandButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    return ctrl


def _add_sheet_button(ws, button_name, cell_range_addr, macro_name):
    """Helper to add a button to a worksheet cell range."""
    cell_range = ws.Range(cell_range_addr)
    left = float(cell_range.Left)
    top = float(cell_range.Top)
    width = float(cell_range.Width)
    height = float(cell_range.Height)

    try:
        # Delete if exists
        ws.Shapes(button_name).Delete()
    except:
        pass

    shp = ws.Shapes.AddShape(5, left, top, width, height)  # 5 = msoShapeRoundedRectangle
    shp.Name = button_name
    
    caption = cell_range.Cells(1, 1).Value
    
    shp.TextFrame.Characters().Text = caption
    shp.TextFrame.Characters().Font.Size = 11
    shp.TextFrame.Characters().Font.Bold = True
    shp.TextFrame.Characters().Font.Color = 16777215  # White
    shp.TextFrame.HorizontalAlignment = -4108  # xlCenter
    shp.TextFrame.VerticalAlignment = -4108
    shp.Fill.ForeColor.RGB = 7884319  # Dark blue
    shp.Line.Visible = False
    shp.OnAction = macro_name



# ═══════════════════════════════════════════════════════════════════════════════
# NAVIGATION BUTTONS ON CONTROL SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def create_nav_buttons(wb):
    """Add navigation shape-buttons to the Control sheet."""
    ws = wb.Sheets("Control")

    # Remove placeholder text (now the cells themselves will have the text)
    for row in [9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35]:  # All button rows
        ws.Range(f"A{row}:C{row}").ClearContents

    # Set cell values which will become button captions
    ws.Range("A9").Value = "Daily Bed Entry"
    ws.Range("A11").Value = "Record Admission"
    ws.Range("A13").Value = "Record Death"
    ws.Range("A15").Value = "Record Ages Entry"
    ws.Range("A17").Value = "Refresh Reports"
    ws.Range("A19").Value = "Manage Wards"  # New button
    ws.Range("A21").Value = "Export Ward Config"  # New button
    ws.Range("A23").Value = "Export Year-End"  # Moved down

    _add_sheet_button(ws, "btnDailyEntry", "Control!A9:C9", "ShowDailyEntry")
    _add_sheet_button(ws, "btnAdmission", "Control!A11:C11", "ShowAdmission")
    _add_sheet_button(ws, "btnDeath", "Control!A13:C13", "ShowDeath")
    _add_sheet_button(ws, "btnAgesEntry", "Control!A15:C15", "ShowAgesEntry")
    _add_sheet_button(ws, "btnRefresh", "Control!A17:C17", "ShowRefreshReports")
    _add_sheet_button(ws, "btnManageWards", "Control!A19:C19", "ShowWardManager")  # New
    _add_sheet_button(ws, "btnExportConfig", "Control!A21:C21", "ExportWardConfig")  # New
    _add_sheet_button(ws, "btnExportYearEnd", "Control!A23:C23", "ExportCarryForward")  # Moved
    _add_sheet_button(ws, "btnPreferences", "Control!A25:C25", "ShowPreferencesInfo")  # New

    # Rebuild button (special orange button)
    _add_sheet_button(ws, "btnRebuild", "Control!A27:C27", "RebuildWorkbookWithPreferences")

    # Diagnostic buttons (row 29, 31, 33 for spacing)
    ws.Range("A29").Value = "Import from Old Workbook"
    ws.Range("A31").Value = "Recalculate All Data"
    ws.Range("A33").Value = "Verify Calculations"
    _add_sheet_button(ws, "btnImport", "Control!A29:C29", "ImportFromOldWorkbook")
    _add_sheet_button(ws, "btnRecalcAll", "Control!A31:C31", "RecalculateAllRows")
    _add_sheet_button(ws, "btnVerify", "Control!A33:C33", "VerifyCalculations")
    # Note: "Fix Date Formats" button removed - date formats now initialized automatically during build


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN INJECTION FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def initialize_date_formats(wb):
    """
    Initialize date column formats for all data tables.
    This ensures date columns are properly formatted from the start.
    """
    # DailyData - EntryDate (col A) and EntryTimestamp (col L)
    try:
        daily_tbl = wb.Sheets("DailyData").ListObjects("tblDaily")
        daily_tbl.ListColumns(1).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        daily_tbl.ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format DailyData columns: {e}")

    # Admissions - AdmissionDate (col B) and EntryTimestamp (col K)
    try:
        adm_tbl = wb.Sheets("Admissions").ListObjects("tblAdmissions")
        adm_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        adm_tbl.ListColumns(11).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format Admissions columns: {e}")

    # DeathsData - DateOfDeath (col B) and EntryTimestamp (col M)
    try:
        deaths_tbl = wb.Sheets("DeathsData").ListObjects("tblDeaths")
        deaths_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        deaths_tbl.ListColumns(13).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format DeathsData columns: {e}")

    # TransfersData - TransferDate (col B) and EntryTimestamp (col H)
    try:
        trans_tbl = wb.Sheets("TransfersData").ListObjects("tblTransfers")
        trans_tbl.ListColumns(2).DataBodyRange.NumberFormat = "yyyy-mm-dd"
        trans_tbl.ListColumns(8).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
    except Exception as e:
        print(f"    Warning: Could not format TransfersData columns: {e}")


def inject_vba(xlsx_path: str, xlsm_path: str, config: WorkbookConfig):
    """Open xlsx in Excel via COM, inject VBA, save as xlsm."""
    import win32com.client

    abs_xlsx = os.path.abspath(xlsx_path)
    abs_xlsm = os.path.abspath(xlsm_path)

    # Verify xlsx file exists
    if not os.path.exists(abs_xlsx):
        raise FileNotFoundError(f"Cannot find xlsx file: {abs_xlsx}")

    print(f"Opening file: {abs_xlsx}")

    # Remove existing xlsm if it exists
    if os.path.exists(abs_xlsm):
        print(f"Removing existing xlsm: {abs_xlsm}")
        os.remove(abs_xlsm)

    # Pre-flight checks
    print("Performing pre-flight checks...")
    
    # Check if file is already open in Excel
    try:
        import psutil
        excel_processes = [p for p in psutil.process_iter(['name', 'open_files']) if p.info['name'] and 'excel' in p.info['name'].lower()]
        for proc in excel_processes:
            try:
                if proc.info['open_files']:
                    for file in proc.info['open_files']:
                        if abs_xlsx in file.path:
                            print(f"\n⚠️  WARNING: File is already open in Excel (PID: {proc.pid})")
                            print("Please close the file in Excel and try again.")
                            raise RuntimeError(f"File is already open: {abs_xlsx}")
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                pass
    except ImportError:
        # psutil not available, skip this check
        print("  (psutil not available, skipping open file check)")
    
    print("Starting Excel for VBA injection...")
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Give Excel a moment to fully initialize
        time.sleep(1)
        
        # Use standard Windows paths (backslashes) which are reliable for local files
        xlsx_path_normalized = abs_xlsx
        print(f"Opening workbook: {xlsx_path_normalized}")
        
        # Try opening with retry logic (sometimes COM needs a moment)
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                # Open with explicit parameters to avoid issues
                wb = excel.Workbooks.Open(
                    xlsx_path_normalized,
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False
                )
                print("  [OK] Workbook opened successfully")
                break
            except Exception as open_error:
                if attempt < max_retries - 1:
                    print(f"  Attempt {attempt + 1} failed, retrying in {retry_delay}s...")
                    time.sleep(retry_delay)
                else:
                    raise open_error
        
        if wb is None:
            raise RuntimeError("Failed to open workbook after all retries")
        vbproj = wb.VBProject

        # 1. Inject standard modules
        print("  Injecting VBA modules...")
        modules = [
            ("modConfig", VBA_MOD_CONFIG),
            ("modDataAccess", VBA_MOD_DATA_ACCESS),
            ("modReports", VBA_MOD_REPORTS),
            ("modNavigation", VBA_MOD_NAVIGATION),
            ("modYearEnd", VBA_MOD_YEAREND),
        ]
        for mod_name, mod_code in modules:
            module = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
            module.Name = mod_name
            module.CodeModule.AddFromString(mod_code)

        # 2. Inject ThisWorkbook code
        print("  Injecting ThisWorkbook code...")
        tb = vbproj.VBComponents("ThisWorkbook")
        tb.CodeModule.AddFromString(VBA_THISWORKBOOK)

        # 2.5. Inject DailyData worksheet change event
        print("  Injecting DailyData worksheet event...")
        daily_data_injected = False

        for comp in vbproj.VBComponents:
            try:
                # Only check worksheet components (Type 100 = vbext_ct_Document)
                if comp.Type == 100:
                    comp_name = comp.Properties("Name").Value
                    if comp_name == "DailyData":
                        comp.CodeModule.AddFromString(VBA_SHEET_DAILYDATA)
                        print(f"    [OK] Event injected into {comp_name} worksheet")
                        daily_data_injected = True
                        break
            except Exception as e:
                # Log specific errors instead of silent failure
                print(f"    ! Component check failed: {e}")
                continue

        if not daily_data_injected:
            raise ValueError("CRITICAL: Failed to inject Worksheet_Change event into DailyData sheet!")

        # 3. Create UserForms
        print("  Creating UserForms...")
        create_daily_entry_form(vbproj)
        create_admission_form(vbproj)
        create_ages_entry_form(vbproj)
        create_death_form(vbproj)
        create_ward_manager_form(vbproj)
        create_preferences_manager_form(vbproj)

        # 4. Add navigation buttons to Control sheet
        print("  Adding navigation buttons...")
        create_nav_buttons(wb)

        # 5. Hide data sheets
        print("  Hiding data sheets...")
        wb.Sheets("DailyData").Visible = 0       # xlSheetHidden
        wb.Sheets("Admissions").Visible = 0
        wb.Sheets("DeathsData").Visible = 0
        wb.Sheets("TransfersData").Visible = 0

        # Hide individual emergency sheets by default
        try:
            wb.Sheets("Male Emergency").Visible = 0
            wb.Sheets("Female Emergency").Visible = 0
        except:
            pass


        # 5.5. Initialize date column formats
        print("  Initializing date column formats...")
        initialize_date_formats(wb)

        # 6. Save as .xlsm (FileFormat 52)
        print(f"  Saving as {abs_xlsm}...")
        wb.SaveAs(abs_xlsm, FileFormat=52)
        wb.Close(SaveChanges=False)

        print(f"Phase 2 complete: {abs_xlsm}")

    except Exception as e:
        print(f"\nERROR during VBA injection: {e}")
        print(f"Error type: {type(e).__name__}")

        # Provide specific troubleshooting based on error
        if "VBProject" in str(e) or "Programmatic access" in str(e):
            print("\nFIX: Enable VBA project access in Excel:")
            print("  1. Open Excel")
            print("  2. File > Options > Trust Center > Trust Center Settings")
            print("  3. Macro Settings > Check 'Trust access to the VBA project object model'")
            print("  4. Click OK and restart Excel")
        elif "Open" in str(e):
            print("\nPossible causes:")
            print(f"  - File exists: {os.path.exists(abs_xlsx)}")
            print(f"  - File path: {abs_xlsx}")
            print(f"  - File readable: {os.access(abs_xlsx, os.R_OK)}")
            print("\nTroubleshooting:")
            print("  1. Try opening the .xlsx file manually in Excel first")
            print("  2. Check if file is corrupted")
            print("  3. Close all Excel windows and try again")
            print("  4. Check antivirus isn't blocking file access")

        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        raise
    finally:
        excel.Quit()
        time.sleep(1)
