Option Explicit

Private Sub UserForm_Initialize()
    ' Load current values from table
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))
        Dim val As Boolean
        val = CBool(tbl.ListRows(i).Range(1, 2).Value)

        If key = "show_emergency_total_remaining" Then
            chkShowEmergencyRemaining.Value = val
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            chkSubtractDeaths.Value = val
        End If
    Next i
End Sub

Private Sub SavePreferencesToTable()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

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
End Sub

Private Sub btnSave_Click()
    SavePreferencesToTable

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
    SavePreferencesToTable

    ' Close form
    Unload Me

    ' Trigger automated rebuild
    RebuildWorkbookWithPreferences
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
