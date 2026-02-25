'###################################################################
'# MODULE: modReports
'# PURPOSE: Generate and refresh statistical reports
'###################################################################

Option Explicit

'===================================================================
' REPORT GENERATION FUNCTIONS
'===================================================================

Public Sub RefreshDeathsReport()
    ' Deaths Summary is now formula-driven and updates automatically
    ' This function forces recalculation if needed
    On Error Resume Next
    Application.ScreenUpdating = False

    Dim wsReport As Worksheet
    Set wsReport = ThisWorkbook.Sheets("Deaths Summary")

    If Not wsReport Is Nothing Then
        wsReport.Calculate
    End If

    Application.ScreenUpdating = True
    MsgBox "Deaths Summary updated (formula-driven report)", vbInformation
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
