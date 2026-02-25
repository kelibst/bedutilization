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
Public Const COL_DEATH_AGE_UNIT As Integer = 8    ' Moved from 12 to align with table structure
Public Const COL_DEATH_SEX As Integer = 9         ' Moved from 8
Public Const COL_DEATH_NHIS As Integer = 10       ' Moved from 9
Public Const COL_DEATH_CAUSE As Integer = 11      ' Moved from 10
Public Const COL_DEATH_WITHIN_24HR As Integer = 12 ' Moved from 11
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
    ' prefix: Optional prefix (e.g., "D" for deaths, "A" for admissions)
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

    ' Find the maximum ID number across ALL rows (not just last row)
    Dim maxNum As Long
    maxNum = 0
    Dim i As Long
    Dim currentID As String
    Dim dashPos As Long
    Dim currentNum As Long

    For i = 1 To tbl.ListRows.Count
        currentID = CStr(tbl.ListRows(i).Range(1, 1).Value)

        ' Only process IDs that match our prefix and year
        If Left(currentID, Len(prefix & yr & "-")) = prefix & yr & "-" Then
            dashPos = InStr(currentID, "-")
            If dashPos > 0 Then
                On Error Resume Next
                currentNum = CLng(Mid(currentID, dashPos + 1))
                If Err.Number = 0 And currentNum > maxNum Then
                    maxNum = currentNum
                End If
                On Error GoTo 0
            End If
        End If
    Next i

    ' Generate next ID
    GenerateNextID = prefix & yr & "-" & Format(maxNum + 1, "00000")
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
    Dim importErrors As Long
    importErrors = 0
    Dim errorLog As String
    
    For i = 1 To oldTbl.ListRows.Count
        ' Error handling for individual rows
        On Error Resume Next
        
        ' Check if row has data
        Dim oldDate As Variant
        oldDate = oldTbl.ListRows(i).Range(1, 1).Value
        
        ' Skip if cell has error or is not a date
        Dim isValidDate As Boolean
        isValidDate = False
        
        If Not IsError(oldDate) Then
            If IsDate(oldDate) And oldDate <> "" Then
                isValidDate = True
            End If
        End If

        If isValidDate Then
            ' Determine target row
            Dim newRow As ListRow
            If useSeedRow And importCount = 0 Then
                Set newRow = newTbl.ListRows(1)
            Else
                Set newRow = newTbl.ListRows.Add
            End If

            ' Copy INPUT columns only (1-9) with safety checks
            With newRow.Range
                .Cells(1, 1).Value = oldDate ' EntryDate
                
                ' Helper to safely copy values
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

            If Err.Number <> 0 Then
                importErrors = importErrors + 1
                If importErrors <= 5 Then
                    errorLog = errorLog & "Row " & i & ": " & Err.Description & vbCrLf
                End If
                Err.Clear
            Else
                importCount = importCount + 1
            End If
        End If
        On Error GoTo ErrorHandler ' Restore main handler
    Next i

    ' Import individual death records BEFORE closing workbook
    Dim deathsResult As String
    deathsResult = ImportIndividualRecords(oldWB, "tblDeaths", "DeathsData", _
                                          "tblDeaths", "DeathsData", 13, COL_DEATH_ID)

    ' Import individual admission records BEFORE closing workbook
    Dim admissionsResult As String
    admissionsResult = ImportIndividualRecords(oldWB, "tblAdmissions", "Admissions", _
                                              "tblAdmissions", "Admissions", 11, COL_ADM_ID)

    ' Close old workbook NOW that we're done importing
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

    Dim msg As String
    msg = "IMPORT COMPLETE!" & vbCrLf & vbCrLf & _
          "Daily Data: " & importCount & " records" & vbCrLf & _
          deathsResult & vbCrLf & _
          admissionsResult & vbCrLf & vbCrLf & _
          "All calculations have been recalculated automatically."

    If importErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & "WARNING: " & importErrors & " daily data rows failed to import." & vbCrLf & _
              "First 5 errors:" & vbCrLf & errorLog
    End If

    MsgBox msg, vbInformation, "Import Complete"
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
' HELPER: Import Individual Records (Deaths, Admissions, etc.)
'===================================================================
Private Function ImportIndividualRecords(oldWB As Workbook, _
    oldTableName As String, oldSheetName As String, _
    newTableName As String, newSheetName As String, _
    columnCount As Integer, idColumn As Integer) As String

    On Error GoTo ErrorHandler

    ' Result message with import stats
    Dim resultMsg As String
    resultMsg = ""

    ' Validate old workbook has the table
    Dim oldWS As Worksheet
    On Error Resume Next
    Set oldWS = oldWB.Sheets(oldSheetName)
    On Error GoTo ErrorHandler

    If oldWS Is Nothing Then
        resultMsg = oldTableName & ": Sheet not found (skipped)"
        ImportIndividualRecords = resultMsg
        Exit Function
    End If

    Dim oldTbl As ListObject
    On Error Resume Next
    Set oldTbl = oldWS.ListObjects(oldTableName)
    On Error GoTo ErrorHandler

    If oldTbl Is Nothing Then
        resultMsg = oldTableName & ": Table not found (skipped)"
        ImportIndividualRecords = resultMsg
        Exit Function
    End If

    ' Get new workbook table
    Dim newWS As Worksheet
    Set newWS = ThisWorkbook.Sheets(newSheetName)
    Dim newTbl As ListObject
    Set newTbl = newWS.ListObjects(newTableName)

    ' Track stats
    Dim importedCount As Integer
    Dim duplicateCount As Integer
    Dim errorCount As Integer
    importedCount = 0
    duplicateCount = 0
    errorCount = 0

    ' Build dictionary of existing IDs to detect duplicates
    Dim existingIDs As Object
    Set existingIDs = CreateObject("Scripting.Dictionary")

    Dim j As Long
    For j = 1 To newTbl.ListRows.Count
        Dim existingID As String
        existingID = CStr(newTbl.ListRows(j).Range(1, idColumn).Value)
        If existingID <> "" Then
            existingIDs(existingID) = True
        End If
    Next j

    ' Import rows from old table
    Dim i As Long
    For i = 1 To oldTbl.ListRows.Count
        On Error Resume Next

        ' Check if row is empty (skip seed rows)
        Dim firstCell As Variant
        firstCell = oldTbl.ListRows(i).Range(1, 1).Value
        If IsEmpty(firstCell) Or firstCell = "" Then
            GoTo NextRow
        End If

        ' Get original ID
        Dim originalID As String
        originalID = CStr(oldTbl.ListRows(i).Range(1, idColumn).Value)

        ' Handle duplicate ID by generating new one with -IMP suffix
        Dim finalID As String
        finalID = originalID
        Dim suffix As Integer
        suffix = 1

        While existingIDs.Exists(finalID)
            If suffix = 1 Then
                finalID = originalID & "-IMP"
            Else
                finalID = originalID & "-IMP" & suffix
            End If
            suffix = suffix + 1
            duplicateCount = duplicateCount + 1
        Wend

        ' Add new ID to dictionary
        existingIDs(finalID) = True

        ' Get row to use (seed row if empty, otherwise new row)
        Dim newRow As ListRow
        If newTbl.ListRows.Count = 1 And _
           (IsEmpty(newTbl.ListRows(1).Range(1, 1).Value) Or _
            newTbl.ListRows(1).Range(1, 1).Value = "") Then
            Set newRow = newTbl.ListRows(1)
        Else
            Set newRow = newTbl.ListRows.Add
        End If

        ' Copy all columns
        Dim col As Integer
        For col = 1 To columnCount
            If col = idColumn Then
                ' Use potentially renamed ID
                newRow.Range(1, col).Value = finalID
            Else
                ' Copy value as-is
                newRow.Range(1, col).Value = oldTbl.ListRows(i).Range(1, col).Value
            End If
        Next col

        importedCount = importedCount + 1

NextRow:
        On Error GoTo ErrorHandler
    Next i

    ' Build result message
    resultMsg = oldTableName & ": " & importedCount & " records"
    If duplicateCount > 0 Then
        resultMsg = resultMsg & " (" & duplicateCount & " IDs renamed)"
    End If

    ImportIndividualRecords = resultMsg
    Exit Function

ErrorHandler:
    errorCount = errorCount + 1
    resultMsg = oldTableName & ": " & importedCount & " records (" & errorCount & " errors)"
    ImportIndividualRecords = resultMsg
End Function

'===================================================================
' DATA SAVE OPERATIONS
'===================================================================

Public Sub SaveAdmission(admDate As Variant, wardCode As String, _
    patientID As String, patientName As String, _
    age As Long, ageUnit As String, sex As String, nhis As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Admissions")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblAdmissions")

    ' Get row to use (seed row if empty, otherwise new row)
    Dim targetRow As ListRow
    Set targetRow = GetOrAddTableRow(tbl)

    ' Generate ID with "A" prefix for Admissions
    Dim newID As String
    newID = GenerateNextID(tbl, "A")

    With targetRow.Range
        .Cells(1, COL_ADM_ID).Value = newID
        
        ' Handle date - try to convert if acceptable, otherwise leave as is (which might fail formatting but saves data)
        If IsDate(admDate) Then
            .Cells(1, COL_ADM_DATE).Value = CDate(admDate)
        Else
            .Cells(1, COL_ADM_DATE).Value = admDate
        End If

        ' IMPORTANT: Date columns must be formatted explicitly
        ' Excel may store dates as text if format isn't set
        ' See initialize_date_formats() in phase2_vba.py for column setup
        .Cells(1, COL_ADM_DATE).NumberFormat = "yyyy-mm-dd"

        If IsDate(admDate) Then
            .Cells(1, COL_ADM_MONTH).Value = Month(CDate(admDate))
        Else
            .Cells(1, COL_ADM_MONTH).Value = 0
        End If
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

Public Sub SaveDeath(deathDate As Variant, wardCode As String, _
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
        
        If IsDate(deathDate) Then
            .Cells(1, COL_DEATH_DATE).Value = CDate(deathDate)
        Else
            .Cells(1, COL_DEATH_DATE).Value = deathDate
        End If

        ' IMPORTANT: Date columns must be formatted explicitly
        ' Excel may store dates as text if format isn't set
        ' See initialize_date_formats() in phase2_vba.py for column setup
        .Cells(1, COL_DEATH_DATE).NumberFormat = "yyyy-mm-dd"

        If IsDate(deathDate) Then
            .Cells(1, COL_DEATH_MONTH).Value = Month(CDate(deathDate))
        Else
            .Cells(1, COL_DEATH_MONTH).Value = 0
        End If
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
