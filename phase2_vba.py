"""
Phase 2: Inject VBA code into the workbook using win32com
Creates UserForms, standard modules, navigation buttons, and saves as .xlsm
"""
import os
import time
import json
from config import WorkbookConfig

# ═══════════════════════════════════════════════════════════════════════════════
# VBA SOURCE CODE - STANDARD MODULES
# ═══════════════════════════════════════════════════════════════════════════════

VBA_MOD_CONFIG = '''
Option Explicit

Public Const HOSPITAL_NAME As String = "HOHOE MUNICIPAL HOSPITAL"
Public Const NUM_WARDS As Long = 9

Public Function GetWardCodes() As Variant
    GetWardCodes = Array("MW", "FW", "CW", "BF", "BG", "BH", "NICU", "MAE", "FAE")
End Function

Public Function GetWardNames() As Variant
    GetWardNames = Array("Male Medical", "Female Medical", "Paediatric", _
        "Block F", "Block G", "Block H", "Neonatal", "Male Emergency", "Female Emergency")
End Function

Public Function GetReportYear() As Long
    GetReportYear = ThisWorkbook.Sheets("Control").Range("B5").Value
End Function
'''

VBA_MOD_DATA_ACCESS = '''
Option Explicit

Public Function GetLastRemainingForWard(wardCode As String, beforeDate As Date) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    Dim bestDate As Date
    Dim bestRemaining As Long
    bestDate = #1/1/1900#
    bestRemaining = 0

    ' If table has no real data, check config for PrevYearRemaining
    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            bestRemaining = GetPrevYearRemaining(wardCode)
            GetLastRemainingForWard = bestRemaining
            Exit Function
        End If
    End If

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim rowDate As Variant
        rowDate = tbl.ListRows(i).Range(1, 1).Value
        If IsDate(rowDate) Then
            If CDate(rowDate) < beforeDate And _
               tbl.ListRows(i).Range(1, 3).Value = wardCode Then
                If CDate(rowDate) > bestDate Then
                    bestDate = CDate(rowDate)
                    bestRemaining = CLng(tbl.ListRows(i).Range(1, 11).Value)
                End If
            End If
        End If
    Next i

    ' If no previous entry found, use PrevYearRemaining
    If bestDate = #1/1/1900# Then
        bestRemaining = GetPrevYearRemaining(wardCode)
    End If

    GetLastRemainingForWard = bestRemaining
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
        rowDate = tbl.ListRows(i).Range(1, 1).Value
        If IsDate(rowDate) Then
            If CDate(rowDate) = entryDate And _
               tbl.ListRows(i).Range(1, 3).Value = wardCode Then
                CheckDuplicateDaily = i
                Exit Function
            End If
        End If
    Next i
    CheckDuplicateDaily = 0
End Function

Public Sub SaveDailyEntry(entryDate As Date, wardCode As String, _
    admissions As Long, discharges As Long, deaths As Long, _
    deathsU24 As Long, transIn As Long, transOut As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDaily")

    Dim prevRemaining As Long
    prevRemaining = GetLastRemainingForWard(wardCode, entryDate)

    ' Check if this is the seed row (empty first row)
    Dim useSeedRow As Boolean
    useSeedRow = False
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
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

    With targetRow.Range
        .Cells(1, 1).Value = entryDate
        .Cells(1, 1).NumberFormat = "yyyy-mm-dd"
        .Cells(1, 2).Value = Month(entryDate)
        .Cells(1, 3).Value = wardCode
        .Cells(1, 4).Value = admissions
        .Cells(1, 5).Value = discharges
        .Cells(1, 6).Value = deaths
        .Cells(1, 7).Value = deathsU24
        .Cells(1, 8).Value = transIn
        .Cells(1, 9).Value = transOut
        .Cells(1, 10).Value = prevRemaining
        ' Column 11 (Remaining) is a formula - auto-calculated
        .Cells(1, 12).Value = Now
        .Cells(1, 12).NumberFormat = "yyyy-mm-dd hh:mm"
    End With
End Sub

Public Sub SaveAdmission(admDate As Date, wardCode As String, _
    patientID As String, patientName As String, _
    age As Long, ageUnit As String, sex As String, nhis As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Admissions")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblAdmissions")

    ' Check seed row
    Dim useSeedRow As Boolean
    useSeedRow = False
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            useSeedRow = True
        End If
    End If

    Dim targetRow As ListRow
    If useSeedRow Then
        Set targetRow = tbl.ListRows(1)
    Else
        Set targetRow = tbl.ListRows.Add
    End If

    ' Generate ID
    Dim newID As String
    newID = GenerateAdmissionID()

    With targetRow.Range
        .Cells(1, 1).Value = newID
        .Cells(1, 2).Value = admDate
        .Cells(1, 2).NumberFormat = "yyyy-mm-dd"
        .Cells(1, 3).Value = Month(admDate)
        .Cells(1, 4).Value = wardCode
        .Cells(1, 5).Value = patientID
        .Cells(1, 6).Value = patientName
        .Cells(1, 7).Value = age
        .Cells(1, 8).Value = ageUnit
        .Cells(1, 9).Value = sex
        .Cells(1, 10).Value = nhis
        .Cells(1, 11).Value = Now
        .Cells(1, 11).NumberFormat = "yyyy-mm-dd hh:mm"
    End With
End Sub

Private Function GenerateAdmissionID() As String
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Admissions").ListObjects("tblAdmissions")
    Dim yr As Long
    yr = GetReportYear()

    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            GenerateAdmissionID = yr & "-00001"
            Exit Function
        End If
    End If

    Dim lastID As String
    lastID = CStr(tbl.ListRows(tbl.ListRows.Count).Range(1, 1).Value)
    Dim dashPos As Long
    dashPos = InStr(lastID, "-")
    If dashPos > 0 Then
        Dim num As Long
        num = CLng(Mid(lastID, dashPos + 1)) + 1
        GenerateAdmissionID = yr & "-" & Format(num, "00000")
    Else
        GenerateAdmissionID = yr & "-00001"
    End If
End Function

Public Sub SaveDeath(deathDate As Date, wardCode As String, _
    folderNum As String, deceasedName As String, _
    age As Long, ageUnit As String, sex As String, _
    nhis As String, causeOfDeath As String, within24 As Boolean)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DeathsData")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblDeaths")

    Dim useSeedRow As Boolean
    useSeedRow = False
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            useSeedRow = True
        End If
    End If

    Dim targetRow As ListRow
    If useSeedRow Then
        Set targetRow = tbl.ListRows(1)
    Else
        Set targetRow = tbl.ListRows.Add
    End If

    ' Generate ID
    Dim newID As String
    newID = GenerateDeathID()

    With targetRow.Range
        .Cells(1, 1).Value = newID
        .Cells(1, 2).Value = deathDate
        .Cells(1, 2).NumberFormat = "yyyy-mm-dd"
        .Cells(1, 3).Value = Month(deathDate)
        .Cells(1, 4).Value = wardCode
        .Cells(1, 5).Value = folderNum
        .Cells(1, 6).Value = deceasedName
        .Cells(1, 7).Value = age
        .Cells(1, 8).Value = ageUnit
        .Cells(1, 9).Value = sex
        .Cells(1, 10).Value = nhis
        .Cells(1, 11).Value = causeOfDeath
        .Cells(1, 12).Value = within24
        .Cells(1, 13).Value = Now
        .Cells(1, 13).NumberFormat = "yyyy-mm-dd hh:mm"
    End With
End Sub

Private Function GenerateDeathID() As String
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")
    Dim yr As Long
    yr = GetReportYear()

    If tbl.ListRows.Count <= 1 Then
        If IsEmpty(tbl.ListRows(1).Range(1, 1).Value) Or tbl.ListRows(1).Range(1, 1).Value = "" Then
            GenerateDeathID = "D" & yr & "-00001"
            Exit Function
        End If
    End If

    Dim lastID As String
    lastID = CStr(tbl.ListRows(tbl.ListRows.Count).Range(1, 1).Value)
    Dim dashPos As Long
    dashPos = InStr(lastID, "-")
    If dashPos > 0 Then
        Dim num As Long
        num = CLng(Mid(lastID, dashPos + 1)) + 1
        GenerateDeathID = "D" & yr & "-" & Format(num, "00000")
    Else
        GenerateDeathID = "D" & yr & "-00001"
    End If
End Function
'''

VBA_MOD_REPORTS = '''
Option Explicit

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

Public Sub RefreshAllReports()
    RefreshDeathsReport
    RefreshCODSummary
    MsgBox "All reports have been refreshed.", vbInformation, "Reports Updated"
End Sub
'''

VBA_MOD_NAVIGATION = '''
Option Explicit

Public Sub ShowDailyEntry()
    frmDailyEntry.Show
End Sub

Public Sub ShowAdmission()
    frmAdmission.Show
End Sub

Public Sub ShowDeath()
    frmDeath.Show
End Sub

Public Sub ShowRefreshReports()
    RefreshAllReports
End Sub
'''

VBA_MOD_YEAREND = '''
Option Explicit

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

    ' Save to file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\carry_forward_" & yr & ".json"

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonStr
    Close #fNum

    MsgBox "Carry-forward data exported to:" & vbCrLf & filePath, vbInformation, "Year-End Export"
End Sub
'''

# ─── UserForm Code ───────────────────────────────────────────────────────────

VBA_FRM_DAILY_ENTRY_CODE = '''
Option Explicit

Private wardCodes As Variant
Private wardNames As Variant

Private Sub UserForm_Initialize()
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    ' Populate ward combo
    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i

    ' Default date to today
    txtDate.Value = Format(Date, "dd/mm/yyyy")

    ' Select first ward
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    ' Initialize numeric fields to 0
    txtAdmissions.Value = "0"
    txtDischarges.Value = "0"
    txtDeaths.Value = "0"
    txtDeaths24.Value = "0"
    txtTransIn.Value = "0"
    txtTransOut.Value = "0"

    UpdatePrevRemaining
End Sub

Private Sub cmbWard_Change()
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub txtDate_AfterUpdate()
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub UpdatePrevRemaining()
    On Error Resume Next
    If cmbWard.ListIndex < 0 Then Exit Sub

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    Dim entryDate As Date
    entryDate = ParseDate(txtDate.Value)
    If entryDate = #1/1/1900# Then Exit Sub

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
    entryDate = ParseDate(txtDate.Value)
    If entryDate = #1/1/1900# Then Exit Sub

    Dim existRow As Long
    existRow = CheckDuplicateDaily(entryDate, wc)

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
        lblStatus.Caption = "* Existing entry loaded - will update on save"
        lblStatus.ForeColor = RGB(200, 100, 0)
    Else
        txtAdmissions.Value = "0"
        txtDischarges.Value = "0"
        txtDeaths.Value = "0"
        txtDeaths24.Value = "0"
        txtTransIn.Value = "0"
        txtTransOut.Value = "0"
        lblStatus.Caption = ""
    End If
    CalculateRemaining
End Sub

Private Sub txtAdmissions_Change()
    CalculateRemaining
End Sub
Private Sub txtDischarges_Change()
    CalculateRemaining
End Sub
Private Sub txtDeaths_Change()
    CalculateRemaining
End Sub
Private Sub txtTransIn_Change()
    CalculateRemaining
End Sub
Private Sub txtTransOut_Change()
    CalculateRemaining
End Sub

Private Sub CalculateRemaining()
    On Error Resume Next
    Dim prev As Long, adm As Long, dis As Long
    Dim dth As Long, ti As Long, tOut As Long, remVal As Long

    prev = CLng(Val(lblPrevRemaining.Caption))
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    remVal = prev + adm - dis - dth + ti - tOut
    lblRemaining.Caption = CStr(remVal)

    If remVal < 0 Then
        lblRemaining.ForeColor = RGB(255, 0, 0)
    Else
        lblRemaining.ForeColor = RGB(0, 100, 0)
    End If
End Sub

Private Sub btnPrevDay_Click()
    Dim d As Date
    d = ParseDate(txtDate.Value)
    If d > #1/1/1900# Then
        d = d - 1
        txtDate.Value = Format(d, "dd/mm/yyyy")
        UpdatePrevRemaining
        CheckExistingEntry
    End If
End Sub

Private Sub btnNextDay_Click()
    Dim d As Date
    d = ParseDate(txtDate.Value)
    If d > #1/1/1900# Then
        d = d + 1
        txtDate.Value = Format(d, "dd/mm/yyyy")
        UpdatePrevRemaining
        CheckExistingEntry
    End If
End Sub

Private Sub btnToday_Click()
    txtDate.Value = Format(Date, "dd/mm/yyyy")
    UpdatePrevRemaining
    CheckExistingEntry
End Sub

Private Sub btnSaveNext_Click()
    If SaveCurrentEntry() Then
        ' Move to next ward
        If cmbWard.ListIndex < cmbWard.ListCount - 1 Then
            cmbWard.ListIndex = cmbWard.ListIndex + 1
        Else
            MsgBox "All wards completed for this date!", vbInformation
            cmbWard.ListIndex = 0
        End If
    End If
End Sub

Private Sub btnSaveNextDay_Click()
    If SaveCurrentEntry() Then
        ' Advance date to next day and reset to first ward
        Dim d As Date
        d = ParseDate(txtDate.Value)
        If d > #1/1/1900# Then
            d = d + 1
            txtDate.Value = Format(d, "dd/mm/yyyy")
            cmbWard.ListIndex = 0
            UpdatePrevRemaining
            CheckExistingEntry
        End If
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
    entryDate = ParseDate(txtDate.Value)
    If entryDate = #1/1/1900# Then
        MsgBox "Please enter a valid date (dd/mm/yyyy).", vbExclamation
        Exit Function
    End If

    Dim adm As Long, dis As Long, dth As Long
    Dim d24 As Long, ti As Long, tOut As Long
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    d24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))

    If d24 > dth Then
        MsgBox "Deaths <24Hrs cannot exceed total Deaths.", vbExclamation
        Exit Function
    End If

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    SaveDailyEntry entryDate, wc, adm, dis, dth, d24, ti, tOut

    lblStatus.Caption = "Saved: " & wardNames(cmbWard.ListIndex) & " - " & txtDate.Value
    lblStatus.ForeColor = RGB(0, 128, 0)

    SaveCurrentEntry = True
End Function

Private Function ParseDate(dateStr As String) As Date
    On Error GoTo badDate
    ' Try dd/mm/yyyy format
    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        ParseDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        Exit Function
    End If
    ' Try other formats
    ParseDate = CDate(dateStr)
    Exit Function
badDate:
    ParseDate = #1/1/1900#
End Function
'''

VBA_FRM_ADMISSION_CODE = '''
Option Explicit

Private wardCodes As Variant
Private wardNames As Variant

Private Sub UserForm_Initialize()
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0

    txtDate.Value = Format(Date, "dd/mm/yyyy")
    txtAge.Value = ""
    txtPatientID.Value = ""
    txtPatientName.Value = ""
    optMale.Value = True
    optInsured.Value = True
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

    SaveAdmission admDate, wc, Trim(txtPatientID.Value), _
        Trim(txtPatientName.Value), CLng(txtAge.Value), _
        cmbAgeUnit.Value, sex, nhis

    lblStatus.Caption = "Saved: " & txtPatientName.Value
    lblStatus.ForeColor = RGB(0, 128, 0)

    SaveAdmissionEntry = True
End Function

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
            lstRecent.AddItem tbl.ListRows(i).Range(1, 6).Value & " | " & _
                tbl.ListRows(i).Range(1, 4).Value & " | " & _
                tbl.ListRows(i).Range(1, 9).Value & " | Age: " & _
                tbl.ListRows(i).Range(1, 7).Value
        End If
    Next i
End Sub

Private Function ParseDateAdm(dateStr As String) As Date
    On Error GoTo badDate
    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        ParseDateAdm = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        Exit Function
    End If
    ParseDateAdm = CDate(dateStr)
    Exit Function
badDate:
    ParseDateAdm = #1/1/1900#
End Function
'''

VBA_FRM_DEATH_CODE = '''
Option Explicit

Private wardCodes As Variant
Private wardNames As Variant

Private Sub UserForm_Initialize()
    wardCodes = GetWardCodes()
    wardNames = GetWardNames()

    Dim i As Long
    For i = 0 To UBound(wardNames)
        cmbWard.AddItem wardNames(i)
    Next i
    If cmbWard.ListCount > 0 Then cmbWard.ListIndex = 0

    cmbAgeUnit.AddItem "Years"
    cmbAgeUnit.AddItem "Months"
    cmbAgeUnit.AddItem "Days"
    cmbAgeUnit.ListIndex = 0

    txtDate.Value = Format(Date, "dd/mm/yyyy")
    optMale.Value = True
    optInsured.Value = True
    chkWithin24.Value = False

    ' Populate cause of death combo with previous entries
    PopulateCauses
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
        txtFolderNum.SetFocus
    End If
End Sub

Private Sub btnSaveClose_Click()
    If SaveDeathEntry() Then
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
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

    SaveDeath deathDate, wc, Trim(txtFolderNum.Value), _
        Trim(txtName.Value), CLng(txtAge.Value), _
        cmbAgeUnit.Value, sex, nhis, _
        Trim(cmbCause.Value), chkWithin24.Value

    lblStatus.Caption = "Saved: " & txtName.Value
    lblStatus.ForeColor = RGB(0, 128, 0)

    SaveDeathEntry = True
End Function

Private Function ParseDateDth(dateStr As String) As Date
    On Error GoTo badDate
    Dim parts() As String
    parts = Split(dateStr, "/")
    If UBound(parts) = 2 Then
        ParseDateDth = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        Exit Function
    End If
    ParseDateDth = CDate(dateStr)
    Exit Function
badDate:
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


# ═══════════════════════════════════════════════════════════════════════════════
# USERFORM CREATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def create_daily_entry_form(vbproj):
    """Create the frmDailyEntry UserForm programmatically."""
    form = vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = "frmDailyEntry"
    form.Properties("Caption").Value = "Daily Bed State Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 510

    d = form.Designer

    y = 12  # current Y position

    # Date label, textbox, and navigation buttons
    _add_label(d, "lblDateLabel", "Date (dd/mm/yyyy):", 12, y, 120, 18)
    _add_textbox(d, "txtDate", 140, y, 100, 20)
    _add_button(d, "btnPrevDay", "< Prev", 248, y, 55, 20)
    _add_button(d, "btnNextDay", "Next >", 308, y, 55, 20)
    _add_button(d, "btnToday", "Today", 368, y, 42, 20)
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
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18)
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18)
    y += 28

    # NHIS radio buttons
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18)
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18)
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
    _add_label(d, "lblRecentLabel", "Recent Admissions:", 12, y, 200, 18)
    y += 20
    lst = d.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstRecent"
    lst.Left = 12
    lst.Top = y
    lst.Width = 390
    lst.Height = 100

    form.CodeModule.AddFromString(VBA_FRM_ADMISSION_CODE)


def create_death_form(vbproj):
    """Create the frmDeath UserForm."""
    form = vbproj.VBComponents.Add(3)
    form.Name = "frmDeath"
    form.Properties("Caption").Value = "Death Record Entry"
    form.Properties("Width").Value = 420
    form.Properties("Height").Value = 480

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
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 18)
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 18)
    y += 28

    # NHIS
    _add_label(d, "lblNHISLabel", "NHIS:", 12, y, 60, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 80, 18)
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 170, y, 100, 18)
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

    form.CodeModule.AddFromString(VBA_FRM_DEATH_CODE)


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


def _add_optionbutton(designer, name, caption, left, top, width, height):
    ctrl = designer.Controls.Add("Forms.OptionButton.1")
    ctrl.Name = name
    ctrl.Caption = caption
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


# ═══════════════════════════════════════════════════════════════════════════════
# NAVIGATION BUTTONS ON CONTROL SHEET
# ═══════════════════════════════════════════════════════════════════════════════

def create_nav_buttons(wb):
    """Add navigation shape-buttons to the Control sheet."""
    ws = wb.Sheets("Control")

    # Remove placeholder text
    for row in [9, 11, 13, 15]:
        ws.Range(f"A{row}:C{row}").ClearContents

    buttons = [
        ("Daily Bed Entry",    "ShowDailyEntry",    9),
        ("Record Admission",   "ShowAdmission",     11),
        ("Record Death",       "ShowDeath",         13),
        ("Refresh Reports",    "ShowRefreshReports", 15),
        ("Export Year-End",    "ExportCarryForward", 17),
    ]

    for caption, macro_name, row_num in buttons:
        # Get the position from the cell
        cell_range = ws.Range(f"A{row_num}")
        left = float(cell_range.Left)
        top = float(cell_range.Top)
        width = 200
        height = 30

        shp = ws.Shapes.AddShape(5, left, top, width, height)  # 5 = msoShapeRoundedRectangle
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
# MAIN INJECTION FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

def inject_vba(xlsx_path: str, xlsm_path: str, config: WorkbookConfig):
    """Open xlsx in Excel via COM, inject VBA, save as xlsm."""
    import win32com.client

    abs_xlsx = os.path.abspath(xlsx_path)
    abs_xlsm = os.path.abspath(xlsm_path)

    # Remove existing xlsm if it exists
    if os.path.exists(abs_xlsm):
        os.remove(abs_xlsm)

    print("Starting Excel for VBA injection...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_xlsx)
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

        # 3. Create UserForms
        print("  Creating Daily Entry form...")
        create_daily_entry_form(vbproj)

        print("  Creating Admission form...")
        create_admission_form(vbproj)

        print("  Creating Death form...")
        create_death_form(vbproj)

        # 4. Add navigation buttons to Control sheet
        print("  Adding navigation buttons...")
        create_nav_buttons(wb)

        # 5. Hide data sheets
        print("  Hiding data sheets...")
        wb.Sheets("DailyData").Visible = 0       # xlSheetHidden
        wb.Sheets("Admissions").Visible = 0
        wb.Sheets("DeathsData").Visible = 0
        wb.Sheets("TransfersData").Visible = 0

        # 6. Save as .xlsm (FileFormat 52)
        print(f"  Saving as {abs_xlsm}...")
        wb.SaveAs(abs_xlsm, FileFormat=52)
        wb.Close(SaveChanges=False)

        print(f"Phase 2 complete: {abs_xlsm}")

    except Exception as e:
        print(f"ERROR during VBA injection: {e}")
        print("\nCommon fix: In Excel, go to:")
        print("  File > Options > Trust Center > Trust Center Settings")
        print("  > Macro Settings > Check 'Trust access to the VBA project object model'")
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        raise
    finally:
        excel.Quit()
        time.sleep(1)
