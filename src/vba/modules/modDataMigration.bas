Attribute VB_Name = "modDataMigration"
'===================================================================
' DATA MIGRATION MODULE
'===================================================================
' This module contains utilities for migrating death data from the old
' column format to the new format where AgeUnit is positioned at column 8
' (immediately after Age) instead of column 12.
'===================================================================

Option Explicit

'===================================================================
' PUBLIC MIGRATION INTERFACE
'===================================================================

Public Sub CheckAndMigrateDeathData()
    ' Detects if data is in old (incorrect) format by checking if:
    ' - Column 8 contains M/F (old Sex position)
    ' - Column 12 contains Years/Months/Days (old AgeUnit position)

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "DeathsData table not found.", vbExclamation
        Exit Sub
    End If

    If tbl.ListRows.Count = 0 Then
        MsgBox "No death records found. No migration needed.", vbInformation
        Exit Sub
    End If

    ' Check first row to determine if migration is needed
    Dim col8Val As String, col12Val As String
    col8Val = CStr(tbl.ListRows(1).Range(1, 8).Value)
    col12Val = CStr(tbl.ListRows(1).Range(1, 12).Value)

    ' Valid age units are: "Years", "Months", "Days"
    ' If column 12 has age units and column 8 has sex (M/F), data is in OLD format
    If (col12Val = "Years" Or col12Val = "Months" Or col12Val = "Days") And _
       (col8Val = "M" Or col8Val = "F") Then
        ' Data is in OLD format - needs migration
        Dim msg As String
        msg = "Death data appears to be in the old column format and needs migration." & vbNewLine & vbNewLine & _
              "This will reorder columns 8-12 to match the new structure:" & vbNewLine & _
              "  Column 8: Sex -> AgeUnit" & vbNewLine & _
              "  Column 9: NHIS -> Sex" & vbNewLine & _
              "  Column 10: Cause -> NHIS" & vbNewLine & _
              "  Column 11: Within24 -> Cause" & vbNewLine & _
              "  Column 12: AgeUnit -> Within24" & vbNewLine & vbNewLine & _
              "IMPORTANT: Backup your workbook before proceeding!" & vbNewLine & vbNewLine & _
              "Continue with migration?"

        If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Data Migration Required") = vbYes Then
            MigrateDeathDataColumns
        Else
            MsgBox "Migration cancelled. Your data has not been changed.", vbInformation
        End If
    Else
        MsgBox "Death data is already in correct format. No migration needed.", vbInformation
    End If
End Sub

'===================================================================
' PRIVATE MIGRATION IMPLEMENTATION
'===================================================================

Private Sub MigrateDeathDataColumns()
    ' Reorders columns 8-12 from old format to new format
    On Error GoTo MigrationError

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")

    Dim totalRows As Long
    totalRows = tbl.ListRows.Count

    Dim i As Long
    Dim migratedCount As Long
    migratedCount = 0

    For i = 1 To totalRows
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) Then
            ' Read old positions (save to variables)
            Dim oldSex As Variant, oldNHIS As Variant, oldCause As Variant
            Dim oldWithin24 As Variant, oldAgeUnit As Variant

            oldSex = tbl.ListRows(i).Range(1, 8).Value       ' Was at col 8
            oldNHIS = tbl.ListRows(i).Range(1, 9).Value      ' Was at col 9
            oldCause = tbl.ListRows(i).Range(1, 10).Value    ' Was at col 10
            oldWithin24 = tbl.ListRows(i).Range(1, 11).Value ' Was at col 11
            oldAgeUnit = tbl.ListRows(i).Range(1, 12).Value  ' Was at col 12

            ' Write to new positions
            tbl.ListRows(i).Range(1, 8).Value = oldAgeUnit   ' Now at col 8
            tbl.ListRows(i).Range(1, 9).Value = oldSex       ' Now at col 9
            tbl.ListRows(i).Range(1, 10).Value = oldNHIS     ' Now at col 10
            tbl.ListRows(i).Range(1, 11).Value = oldCause    ' Now at col 11
            tbl.ListRows(i).Range(1, 12).Value = oldWithin24 ' Now at col 12

            migratedCount = migratedCount + 1
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Migration complete!" & vbNewLine & vbNewLine & _
           "Records migrated: " & migratedCount & vbNewLine & _
           "Total rows processed: " & totalRows, vbInformation, "Migration Successful"
    Exit Sub

MigrationError:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during migration: " & Err.Description & vbNewLine & vbNewLine & _
           "Row: " & i & vbNewLine & _
           "Migration aborted. Please check your data.", vbCritical
End Sub

'===================================================================
' VALIDATION UTILITY
'===================================================================

Public Sub ValidateDeathData()
    ' Validates that all death records have correct data in correct columns
    ' Useful for verifying migration success

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("DeathsData").ListObjects("tblDeaths")
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "DeathsData table not found.", vbExclamation
        Exit Sub
    End If

    If tbl.ListRows.Count = 0 Then
        MsgBox "No death records to validate.", vbInformation
        Exit Sub
    End If

    Dim i As Long
    Dim invalidCount As Long
    Dim invalidRows As String
    invalidCount = 0
    invalidRows = ""

    For i = 1 To tbl.ListRows.Count
        If Not IsEmpty(tbl.ListRows(i).Range(1, 1).Value) Then
            Dim ageUnit As String, sex As String, nhis As String
            ageUnit = CStr(tbl.ListRows(i).Range(1, 8).Value)
            sex = CStr(tbl.ListRows(i).Range(1, 9).Value)
            nhis = CStr(tbl.ListRows(i).Range(1, 10).Value)

            ' Validate AgeUnit (column 8)
            If ageUnit <> "Years" And ageUnit <> "Months" And ageUnit <> "Days" Then
                Debug.Print "Invalid AgeUnit in row " & i & ": '" & ageUnit & "'"
                invalidCount = invalidCount + 1
                If Len(invalidRows) < 200 Then
                    invalidRows = invalidRows & i & " (AgeUnit), "
                End If
            End If

            ' Validate Sex (column 9)
            If sex <> "M" And sex <> "F" Then
                Debug.Print "Invalid Sex in row " & i & ": '" & sex & "'"
                invalidCount = invalidCount + 1
                If Len(invalidRows) < 200 Then
                    invalidRows = invalidRows & i & " (Sex), "
                End If
            End If

            ' Validate NHIS (column 10)
            If nhis <> "Insured" And nhis <> "Non-Insured" Then
                Debug.Print "Invalid NHIS in row " & i & ": '" & nhis & "'"
                invalidCount = invalidCount + 1
                If Len(invalidRows) < 200 Then
                    invalidRows = invalidRows & i & " (NHIS), "
                End If
            End If
        End If
    Next i

    If invalidCount = 0 Then
        MsgBox "All death data validated successfully!" & vbNewLine & vbNewLine & _
               "Total records checked: " & tbl.ListRows.Count, vbInformation, "Validation Passed"
    Else
        If Len(invalidRows) > 0 Then
            invalidRows = Left(invalidRows, Len(invalidRows) - 2) ' Remove trailing comma
        End If
        MsgBox invalidCount & " validation error(s) found." & vbNewLine & vbNewLine & _
               "Problem rows: " & invalidRows & vbNewLine & vbNewLine & _
               "Check the Immediate window (Ctrl+G) for details.", vbExclamation, "Validation Failed"
    End If
End Sub
