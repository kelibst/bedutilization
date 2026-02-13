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
