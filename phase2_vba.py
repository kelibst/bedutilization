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
    deathsU24 As Long, transIn As Long, transOut As Long, malariaCases As Long)

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

    ' Calculate remaining as a VALUE (not formula - more reliable)
    ' Formula: Remaining = PrevRemaining + Admissions + TransfersIn - (Discharges + Deaths + TransfersOut) - Deaths<24Hrs
    ' Deaths and Deaths<24Hrs are SEPARATE counts (not a subset)
    Dim remaining As Long
    remaining = prevRemaining + admissions + transIn - discharges - deaths - transOut - deathsU24

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
        .Cells(1, 11).Value = remaining
        .Cells(1, 12).Value = malariaCases
        .Cells(1, 13).Value = Now
        .Cells(1, 13).NumberFormat = "yyyy-mm-dd hh:mm"
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

Public Sub ShowAgesEntry()
    frmAgesEntry.Show
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
    txtMalaria.Value = "0"

    isLoading = False
    UpdatePrevRemaining
    CheckExistingEntry
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
        txtMalaria.Value = CStr(tbl.ListRows(existRow).Range(1, 12).Value)
        lblStatus.Caption = "* Existing entry loaded"
        lblStatus.ForeColor = RGB(200, 100, 0)
    Else
        txtAdmissions.Value = "0"
        txtDischarges.Value = "0"
        txtDeaths.Value = "0"
        txtDeaths24.Value = "0"
        txtTransIn.Value = "0"
        txtTransOut.Value = "0"
        txtMalaria.Value = "0"
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
Private Sub txtAdmissions_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtDischarges_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub
Private Sub txtDischarges_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtDeaths_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub
Private Sub txtDeaths_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtDeaths24_Change()
    If Not isLoading Then isDirty = True
    ' Deaths24 doesn't affect Remaining, but trigger Calc anyway
    CalculateRemaining
End Sub
Private Sub txtDeaths24_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtTransIn_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub
Private Sub txtTransIn_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtTransOut_Change()
    If Not isLoading Then isDirty = True
    CalculateRemaining
End Sub
Private Sub txtTransOut_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
End Sub

Private Sub txtMalaria_Change()
    If Not isLoading Then isDirty = True
End Sub
Private Sub txtMalaria_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveCurrentEntry
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
    Dim d24 As Long, ti As Long, tOut As Long, mal As Long
    adm = CLng(Val(txtAdmissions.Value))
    dis = CLng(Val(txtDischarges.Value))
    dth = CLng(Val(txtDeaths.Value))
    d24 = CLng(Val(txtDeaths24.Value))
    ti = CLng(Val(txtTransIn.Value))
    tOut = CLng(Val(txtTransOut.Value))
    mal = CLng(Val(txtMalaria.Value))

    If d24 > dth Then
        MsgBox "Deaths <24Hrs cannot exceed total Deaths.", vbExclamation
        Exit Function
    End If

    Dim wc As String
    wc = wardCodes(cmbWard.ListIndex)

    SaveDailyEntry entryDate, wc, adm, dis, dth, d24, ti, tOut, mal

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

Private Sub UserForm_Initialize()
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

VBA_FRM_AGES_ENTRY_CODE = '''
Option Explicit

Private wardCodes As Variant
Private wardNames As Variant

Private Sub UserForm_Initialize()
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

    Dim yr As Long
    yr = GetReportYear()
    Dim dt As Date
    dt = DateSerial(yr, cmbMonth.ListIndex + 1, spnDay.Value)

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

    ' Save
    Application.Run "SaveAdmission", dt, wc, "-", "Age Entry", age, unit, sex, nhis

    ' Post-Save Reset
    lblStatus.Caption = "Saved: " & age & " " & unit & " (" & sex & ", " & nhis & ")"
    lblStatus.ForeColor = &H8000& ' Green

    txtAge.Value = ""
    cmbAgeUnit.ListIndex = 0 ' Reset to Years
    ' Keep persistent selections (Ward, Date, Sex, NHIS)
    
    txtAge.SetFocus
End Sub

Private Sub btnClose_Click()
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

Private Sub UserForm_Initialize()
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
    form.Properties("Height").Value = 530

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
        ("txtMalaria", "Malaria Cases:"),
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
    form.Properties("Height").Value = 380

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
    _add_optionbutton(d, "optMale", "Male", 80, y, 60, 20)
    _add_optionbutton(d, "optFemale", "Female", 150, y, 70, 20)
    y += 28

    # Insurance
    _add_label(d, "lblIns", "Health Ins:", 12, y, 65, 18)
    _add_optionbutton(d, "optInsured", "Insured", 80, y, 70, 20)
    _add_optionbutton(d, "optNonInsured", "Non-Insured", 160, y, 100, 20)
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

    # Inject
    form.CodeModule.AddFromString(VBA_FRM_AGES_ENTRY_CODE)


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
    for row in [9, 11, 13, 15, 17]: # Added 17 for the new button
        ws.Range(f"A{row}:C{row}").ClearContents

    # Set cell values which will become button captions
    ws.Range("A9").Value = "Daily Bed Entry"
    ws.Range("A11").Value = "Record Admission"
    ws.Range("A13").Value = "Record Death"
    ws.Range("A15").Value = "Record Ages Entry" # New button
    ws.Range("A17").Value = "Refresh Reports" # Moved
    ws.Range("A19").Value = "Export Year-End" # Moved down

    _add_sheet_button(ws, "btnDailyEntry", "Control!A9:C9", "ShowDailyEntry")
    _add_sheet_button(ws, "btnAdmission", "Control!A11:C11", "ShowAdmission")
    _add_sheet_button(ws, "btnDeath", "Control!A13:C13", "ShowDeath")
    _add_sheet_button(ws, "btnAgesEntry", "Control!A15:C15", "ShowAgesEntry")
    _add_sheet_button(ws, "btnRefresh", "Control!A17:C17", "ShowRefreshReports")
    _add_sheet_button(ws, "btnExportYearEnd", "Control!A19:C19", "ExportCarryForward")


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
        print("  Creating UserForms...")
        create_daily_entry_form(vbproj)
        create_admission_form(vbproj)
        create_ages_entry_form(vbproj)
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
