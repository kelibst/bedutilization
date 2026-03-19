Private Sub Workbook_Open()
    ' Navigate to Control sheet
    On Error Resume Next
    ThisWorkbook.Sheets("Control").Activate

    ' Hide data sheets
    ThisWorkbook.Sheets("DailyData").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Admissions").Visible = xlSheetHidden
    ThisWorkbook.Sheets("DeathsData").Visible = xlSheetHidden
    ThisWorkbook.Sheets("TransfersData").Visible = xlSheetHidden

    ' Hide individual emergency sheets by default (added 2026-02)
    ThisWorkbook.Sheets("Male Emergency").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Female Emergency").Visible = xlSheetHidden

    ' Re-apply sheet protection (UserInterfaceOnly is lost on close/reopen)
    ApplySheetProtection
End Sub

'==============================================================================
' APPLY SHEET PROTECTION
' Locks calculated / auto-generated columns on data sheets so users cannot
' accidentally corrupt them.  UserInterfaceOnly:=True lets VBA write freely
' while blocking direct user edits to locked cells.
'
' Ward report sheets (formula-only) are fully locked - users should never
' type into those cells directly.
'
' This sub is called from Workbook_Open so protection is always active.
'==============================================================================
Private Sub ApplySheetProtection()
    On Error Resume Next   ' Silently skip sheets that don't exist yet

    ' ── DailyData ─────────────────────────────────────────────────────────────
    ' Columns: A=EntryDate  B=Month  C=WardCode  D=Admissions  E=Discharges
    '          F=Deaths  G=DeathsUnder24Hrs  H=TransfersIn  I=TransfersOut
    '          J=PrevRemaining (calculated)  K=Remaining (calculated)
    '          L=EntryTimestamp (auto-filled)
    With ThisWorkbook.Sheets("DailyData")
        .Unprotect
        .Cells.Locked = True              ' Lock everything first
        .Range("D:I").Locked = False      ' Unlock user-input columns only
        .Protect Password:="", _
                 UserInterfaceOnly:=True, _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False
    End With

    ' ── Admissions ────────────────────────────────────────────────────────────
    ' Columns: A=AdmissionID (auto)  B=AdmissionDate  C=Month (auto)
    '          D=WardCode  E=PatientID  F=PatientName  G=Age  H=AgeUnit
    '          I=Sex  J=NHIS  K=EntryTimestamp (auto)
    With ThisWorkbook.Sheets("Admissions")
        .Unprotect
        .Cells.Locked = False
        .Range("A:A").Locked = True       ' AdmissionID  - auto-generated
        .Range("C:C").Locked = True       ' Month        - auto-filled
        .Range("K:K").Locked = True       ' Timestamp    - auto-filled
        .Protect Password:="", _
                 UserInterfaceOnly:=True, _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False
    End With

    ' ── DeathsData ────────────────────────────────────────────────────────────
    ' Columns: A=DeathID (auto)  B=DateOfDeath  C=Month (auto)
    '          D=WardCode  E=FolderNumber  F=NameOfDeceased  G=Age  H=AgeUnit
    '          I=Sex  J=NHIS  K=CauseOfDeath  L=DeathWithin24Hrs
    '          M=EntryTimestamp (auto)
    With ThisWorkbook.Sheets("DeathsData")
        .Unprotect
        .Cells.Locked = False
        .Range("A:A").Locked = True       ' DeathID   - auto-generated
        .Range("C:C").Locked = True       ' Month     - auto-filled
        .Range("M:M").Locked = True       ' Timestamp - auto-filled
        .Protect Password:="", _
                 UserInterfaceOnly:=True, _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False
    End With

    ' ── TransfersData ─────────────────────────────────────────────────────────
    ' Columns: A=TransferID (auto)  B=TransferDate  C=Month (auto)
    '          D=FromWardCode  E=ToWardCode  F=PatientID  G=PatientName
    '          H=EntryTimestamp (auto)
    With ThisWorkbook.Sheets("TransfersData")
        .Unprotect
        .Cells.Locked = False
        .Range("A:A").Locked = True       ' TransferID - auto-generated
        .Range("C:C").Locked = True       ' Month      - auto-filled
        .Range("H:H").Locked = True       ' Timestamp  - auto-filled
        .Protect Password:="", _
                 UserInterfaceOnly:=True, _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False
    End With

    ' ── Ward report sheets ────────────────────────────────────────────────────
    ' These sheets are formula-based read-only reports.  Lock all cells so
    ' users cannot accidentally overwrite SUMIFS formulas.
    ' VBA does not write directly to these sheets so UserInterfaceOnly is still
    ' set for consistency.
    Dim ws As Worksheet
    Dim skipList As String
    skipList = ",Control,DailyData,Admissions,DeathsData,TransfersData,"

    For Each ws In ThisWorkbook.Worksheets
        If InStr(skipList, "," & ws.Name & ",") = 0 Then
            ws.Unprotect
            ws.Cells.Locked = True
            ws.Protect Password:="", _
                         UserInterfaceOnly:=True, _
                         DrawingObjects:=False, _
                         Contents:=True, _
                         Scenarios:=False
        End If
    Next ws

    On Error GoTo 0
End Sub
