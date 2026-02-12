"""
Add FixAllDateFormats code to existing workbook
This script opens the existing xlsm file and injects the date fixing VBA code.
"""
import win32com.client
import os
import time

# VBA code for date fixing
FIX_DATE_FORMATS_CODE = '''
Public Sub FixAllDateFormats()
    ' Fix date formatting issues across all data tables
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim fixCount As Long
    fixCount = 0
    Dim report As String
    report = "DATE FORMAT FIX REPORT" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf

    ' Fix DailyData table
    Dim wsDailyData As Worksheet
    Set wsDailyData = ThisWorkbook.Sheets("DailyData")
    Dim tblDaily As ListObject
    Set tblDaily = wsDailyData.ListObjects("tblDaily")

    report = report & "DailyData Table:" & vbCrLf
    fixCount = fixCount + FixDateColumn(tblDaily, 1, "yyyy-mm-dd", "EntryDate")
    fixCount = fixCount + FixDateColumn(tblDaily, 12, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed date cells" & vbCrLf & vbCrLf

    ' Fix Admissions table
    Dim wsAdm As Worksheet
    Set wsAdm = ThisWorkbook.Sheets("Admissions")
    Dim tblAdm As ListObject
    Set tblAdm = wsAdm.ListObjects("tblAdmissions")

    Dim admFixCount As Long
    report = report & "Admissions Table:" & vbCrLf
    admFixCount = FixDateColumn(tblAdm, 2, "yyyy-mm-dd", "AdmissionDate")
    admFixCount = admFixCount + FixDateColumn(tblAdm, 11, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + admFixCount

    ' Fix DeathsData table
    Dim wsDeaths As Worksheet
    Set wsDeaths = ThisWorkbook.Sheets("DeathsData")
    Dim tblDeaths As ListObject
    Set tblDeaths = wsDeaths.ListObjects("tblDeaths")

    Dim deathFixCount As Long
    report = report & "DeathsData Table:" & vbCrLf
    deathFixCount = FixDateColumn(tblDeaths, 2, "yyyy-mm-dd", "DateOfDeath")
    deathFixCount = deathFixCount + FixDateColumn(tblDeaths, 13, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + deathFixCount

    ' Fix TransfersData table
    Dim wsTrans As Worksheet
    Set wsTrans = ThisWorkbook.Sheets("TransfersData")
    Dim tblTrans As ListObject
    Set tblTrans = wsTrans.ListObjects("tblTransfers")

    Dim transFixCount As Long
    report = report & "TransfersData Table:" & vbCrLf
    transFixCount = FixDateColumn(tblTrans, 2, "yyyy-mm-dd", "TransferDate")
    transFixCount = transFixCount + FixDateColumn(tblTrans, 8, "yyyy-mm-dd hh:mm", "EntryTimestamp")
    report = report & "  - Fixed date cells" & vbCrLf & vbCrLf
    fixCount = fixCount + transFixCount

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    report = report & String(50, "=") & vbCrLf
    report = report & "TOTAL: Fixed " & fixCount & " date cells across all tables"

    MsgBox report, vbInformation, "Date Format Fix Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error fixing date formats: " & Err.Description, vbCritical, "Error"
End Sub

Private Function FixDateColumn(tbl As ListObject, colIndex As Long, _
                               dateFormat As String, colName As String) As Long
    Dim fixedCount As Long
    fixedCount = 0

    On Error Resume Next
    tbl.ListColumns(colIndex).DataBodyRange.NumberFormat = dateFormat
    On Error GoTo 0

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim cellVal As Variant
        cellVal = tbl.ListRows(i).Range(1, colIndex).Value

        If Not IsEmpty(cellVal) And cellVal <> "" Then
            If VarType(cellVal) = vbString Then
                If IsDate(cellVal) Then
                    tbl.ListRows(i).Range(1, colIndex).Value = CDate(cellVal)
                    tbl.ListRows(i).Range(1, colIndex).NumberFormat = dateFormat
                    fixedCount = fixedCount + 1
                End If
            ElseIf IsDate(cellVal) Then
                Dim tempDate As Date
                tempDate = CDate(cellVal)
                tbl.ListRows(i).Range(1, colIndex).Value = tempDate
                tbl.ListRows(i).Range(1, colIndex).NumberFormat = dateFormat
                fixedCount = fixedCount + 1
            End If
        End If
    Next i

    FixDateColumn = fixedCount
End Function
'''

def add_date_fix_to_workbook(xlsm_path):
    """Add date fixing code to an existing xlsm workbook"""
    abs_path = os.path.abspath(xlsm_path)

    if not os.path.exists(abs_path):
        print(f"ERROR: File not found: {abs_path}")
        return False

    print(f"Opening workbook: {xlsm_path}")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path)
        vbproj = wb.VBProject

        # Find or create modDataAccess module
        mod_data_access = None
        for component in vbproj.VBComponents:
            if component.Name == "modDataAccess":
                mod_data_access = component
                print("Found existing modDataAccess module")
                break

        if mod_data_access is None:
            print("ERROR: modDataAccess module not found in workbook")
            wb.Close(SaveChanges=False)
            return False

        # Check if FixAllDateFormats already exists
        code_module = mod_data_access.CodeModule
        line_count = code_module.CountOfLines

        existing_code = ""
        if line_count > 0:
            existing_code = code_module.Lines(1, line_count)

        if "Sub FixAllDateFormats" in existing_code:
            print("FixAllDateFormats already exists - replacing...")
            # Find and delete old version
            start_line = 0
            for i in range(1, line_count + 1):
                line = code_module.Lines(i, 1)
                if "Sub FixAllDateFormats" in line:
                    start_line = i
                    break

            if start_line > 0:
                # Find end of procedure
                end_line = start_line
                for i in range(start_line, line_count + 1):
                    line = code_module.Lines(i, 1)
                    if "End Sub" in line and i > start_line:
                        end_line = i
                        break

                # Also delete FixDateColumn function
                for i in range(end_line + 1, line_count + 1):
                    line = code_module.Lines(i, 1)
                    if "Function FixDateColumn" in line:
                        func_start = i
                        for j in range(i, line_count + 1):
                            if "End Function" in code_module.Lines(j, 1):
                                code_module.DeleteLines(func_start, j - func_start + 1)
                                break
                        break

                code_module.DeleteLines(start_line, end_line - start_line + 1)
                print(f"Deleted old code from lines {start_line} to {end_line}")

        # Add new code at the end of the module
        print("Adding FixAllDateFormats code...")
        code_module.AddFromString(FIX_DATE_FORMATS_CODE)

        # Save the workbook
        print("Saving workbook...")
        wb.Save()
        wb.Close(SaveChanges=False)

        print("\n" + "="*60)
        print("SUCCESS! Date fixing code has been added to the workbook.")
        print("="*60)
        print(f"\nYou can now:")
        print(f"1. Open {xlsm_path}")
        print(f"2. Press Alt+F8")
        print(f"3. Select 'FixAllDateFormats'")
        print(f"4. Click 'Run'")
        print("\nThis will fix all date formatting issues in your data!")

        return True

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        excel.Quit()
        time.sleep(1)

if __name__ == "__main__":
    # Add to the working xlsm file
    xlsm_file = "Bed_Utilization_2026new.xlsm"

    if not os.path.exists(xlsm_file):
        print(f"ERROR: {xlsm_file} not found!")
        print("Available xlsm files:")
        for f in os.listdir("."):
            if f.endswith(".xlsm"):
                print(f"  - {f}")
    else:
        add_date_fix_to_workbook(xlsm_file)
