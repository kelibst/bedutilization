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
