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
