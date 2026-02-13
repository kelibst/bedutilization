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
End Sub
