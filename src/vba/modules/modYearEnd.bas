'###################################################################
'# MODULE: modYearEnd
'# PURPOSE: Year-end operations, exports, and workbook rebuild
'###################################################################

Option Explicit

'===================================================================
' YEAR-END EXPORT FUNCTIONS
'===================================================================

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

    ' Save to file using helper function
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\carry_forward_" & yr & ".json"

    Call WriteJSONFile(filePath, jsonStr, "Carry-forward data exported to:")
End Sub

' Export current ward configuration to wards_config.json
Public Sub ExportWardsConfig()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblWardConfig")

    ' Build JSON structure
    Dim jsonStr As String
    jsonStr = "{" & vbCrLf
    jsonStr = jsonStr & "  ""_comment"": ""Ward Configuration for Bed Utilization System""," & vbCrLf
    jsonStr = jsonStr & "  ""_instructions"": [" & vbCrLf
    jsonStr = jsonStr & "    ""To add a new ward, copy an existing ward entry and modify the values.""," & vbCrLf
    jsonStr = jsonStr & "    ""After making changes, save this file and rebuild the workbook.""" & vbCrLf
    jsonStr = jsonStr & "  ]," & vbCrLf
    jsonStr = jsonStr & "  ""wards"": [" & vbCrLf

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        jsonStr = jsonStr & "    {" & vbCrLf
        jsonStr = jsonStr & "      ""code"": """ & tbl.ListRows(i).Range(1, 1).Value & """," & vbCrLf
        jsonStr = jsonStr & "      ""name"": """ & tbl.ListRows(i).Range(1, 2).Value & """," & vbCrLf
        jsonStr = jsonStr & "      ""bed_complement"": " & tbl.ListRows(i).Range(1, 3).Value & "," & vbCrLf
        jsonStr = jsonStr & "      ""prev_year_remaining"": " & tbl.ListRows(i).Range(1, 4).Value & "," & vbCrLf

        Dim isEmerg As Boolean
        isEmerg = tbl.ListRows(i).Range(1, 5).Value
        jsonStr = jsonStr & "      ""is_emergency"": " & LCase(CStr(isEmerg)) & "," & vbCrLf
        jsonStr = jsonStr & "      ""display_order"": " & tbl.ListRows(i).Range(1, 6).Value & vbCrLf
        jsonStr = jsonStr & "    }"
        If i < tbl.ListRows.Count Then jsonStr = jsonStr & ","
        jsonStr = jsonStr & vbCrLf
    Next i

    jsonStr = jsonStr & "  ]" & vbCrLf
    jsonStr = jsonStr & "}" & vbCrLf

    ' Save to wards_config.json
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\wards_config.json"

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonStr
    Close #fNum

    MsgBox "Ward configuration exported to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "To apply changes, rebuild the workbook:" & vbCrLf & _
           "python build_workbook.py --year " & GetReportYear(), _
           vbInformation, "Export Ward Configuration"
End Sub

Public Sub ExportPreferencesConfig()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("Control").ListObjects("tblPreferences")

    ' Read current values from table
    Dim showEmergencyRemaining As Boolean
    Dim subtractDeaths As Boolean

    ' Find rows by key (flexible ordering)
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim key As String
        key = Trim(CStr(tbl.ListRows(i).Range(1, 1).Value))

        If key = "show_emergency_total_remaining" Then
            showEmergencyRemaining = CBool(tbl.ListRows(i).Range(1, 2).Value)
        ElseIf key = "subtract_deaths_under_24hrs_from_admissions" Then
            subtractDeaths = CBool(tbl.ListRows(i).Range(1, 2).Value)
        End If
    Next i

    ' Build JSON structure with proper formatting
    Dim jsonStr As String
    jsonStr = "{" & vbCrLf
    jsonStr = jsonStr & "  ""_comment"": ""Hospital Preferences for Bed Utilization System""," & vbCrLf
    jsonStr = jsonStr & "  ""_instructions"": [" & vbCrLf
    jsonStr = jsonStr & "    ""Configure hospital-specific behavioral preferences here.""," & vbCrLf
    jsonStr = jsonStr & "    ""After making changes, rebuild the workbook using: python build_workbook.py --year YYYY""," & vbCrLf
    jsonStr = jsonStr & "    ""Changing preferences mid-year may cause inconsistency between months.""" & vbCrLf
    jsonStr = jsonStr & "  ]," & vbCrLf
    jsonStr = jsonStr & "  ""version"": ""1.0""," & vbCrLf
    jsonStr = jsonStr & "  ""hospital_name"": ""HOHOE MUNICIPAL HOSPITAL""," & vbCrLf
    jsonStr = jsonStr & "  ""preferences"": {" & vbCrLf
    jsonStr = jsonStr & "    ""show_emergency_total_remaining"": " & LCase(CStr(showEmergencyRemaining)) & "," & vbCrLf
    jsonStr = jsonStr & "    ""subtract_deaths_under_24hrs_from_admissions"": " & LCase(CStr(subtractDeaths)) & vbCrLf
    jsonStr = jsonStr & "  }," & vbCrLf
    jsonStr = jsonStr & "  ""metadata"": {" & vbCrLf
    jsonStr = jsonStr & "    ""created_date"": """ & Format(Date, "yyyy-mm-dd") & """," & vbCrLf
    jsonStr = jsonStr & "    ""description"": ""Hospital-specific configuration preferences for bed utilization tracking""" & vbCrLf
    jsonStr = jsonStr & "  }" & vbCrLf
    jsonStr = jsonStr & "}" & vbCrLf

    ' Save to hospital_preferences.json
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\\config\\hospital_preferences.json"

    On Error GoTo ErrorHandler
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonStr
    Close #fNum

    ' Success message with rebuild instructions
    MsgBox "Hospital preferences exported to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "IMPORTANT: These changes will NOT take effect until you rebuild the workbook:" & vbCrLf & _
           "python build_workbook.py --year " & GetReportYear() & vbCrLf & vbCrLf & _
           "Preferences affect formulas and report structure, so rebuilding is required.", _
           vbInformation, "Export Hospital Preferences"
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting preferences: " & Err.Description, vbCritical, "Export Error"
End Sub

'===================================================================
' JSON HELPER FUNCTIONS
'===================================================================

Private Function WriteJSONFile(filePath As String, jsonContent As String, successMsg As String) As Boolean
    ' Centralized JSON file writing with error handling
    ' Returns True if successful, False otherwise
    On Error GoTo WriteError

    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, jsonContent
    Close #fNum

    If successMsg <> "" Then
        MsgBox successMsg & vbCrLf & filePath, vbInformation, "Export Successful"
    End If

    WriteJSONFile = True
    Exit Function

WriteError:
    WriteJSONFile = False
    MsgBox "Error writing JSON file: " & filePath & vbCrLf & Err.Description, vbCritical, "Export Error"
End Function

'===================================================================
' WORKBOOK REBUILD FUNCTIONS
'===================================================================
' Automated rebuild process with data preservation and validation

Private Function CheckRebuildPrerequisites() As String
    ' Returns empty string if all checks pass, otherwise returns error message
    ' Validates: Python installation, build scripts, config files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if build_workbook.py exists
    Dim buildScript As String
    buildScript = ThisWorkbook.Path & "\\build_workbook.py"
    If Not fso.FileExists(buildScript) Then
        CheckRebuildPrerequisites = "ERROR: build_workbook.py not found in workbook directory." & vbCrLf & _
                                    "Expected: " & buildScript
        Exit Function
    End If

    ' Check if Python is available
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\\python_check.txt"

    On Error Resume Next
    shell.Run "cmd /c python --version > """ & tempFile & """ 2>&1", 0, True
    On Error GoTo 0

    ' Read the result
    If fso.FileExists(tempFile) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(tempFile, 1)
        Dim pythonVersion As String
        If Not ts.AtEndOfStream Then
            pythonVersion = ts.ReadLine
        End If
        ts.Close
        fso.DeleteFile tempFile

        If InStr(LCase(pythonVersion), "python") = 0 Then
            CheckRebuildPrerequisites = "ERROR: Python not found or not in PATH." & vbCrLf & _
                                       "Please install Python and add it to your system PATH."
            Exit Function
        End If
    Else
        CheckRebuildPrerequisites = "ERROR: Unable to check Python installation."
        Exit Function
    End If

    ' Check if config files exist
    Dim wardsConfig As String
    wardsConfig = ThisWorkbook.Path & "\\config\\wards_config.json"
    If Not fso.FileExists(wardsConfig) Then
        CheckRebuildPrerequisites = "ERROR: config\\wards_config.json not found." & vbCrLf & _
                                    "Please export ward configuration first."
        Exit Function
    End If

    ' All checks passed
    CheckRebuildPrerequisites = ""
End Function

Public Sub RebuildWorkbookWithPreferences()
    ' Automated workbook rebuild with data preservation
    On Error GoTo ErrorHandler

    ' Step 1: Check prerequisites
    Dim prereqCheck As String
    prereqCheck = CheckRebuildPrerequisites()
    If prereqCheck <> "" Then
        MsgBox prereqCheck, vbCritical, "Rebuild Prerequisites Failed"
        Exit Sub
    End If

    ' Step 2: Confirm with user
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("This will rebuild the workbook with new preferences." & vbCrLf & vbCrLf & _
                         "Steps:" & vbCrLf & _
                         "1. Current workbook will be backed up" & vbCrLf & _
                         "2. Preferences will be exported to JSON" & vbCrLf & _
                         "3. Workbook will be rebuilt with Python" & vbCrLf & _
                         "4. You can import data from backup" & vbCrLf & vbCrLf & _
                         "Continue?", _
                         vbQuestion + vbYesNo, "Rebuild Workbook")
    If userResponse <> vbYes Then Exit Sub

    ' Step 3: Export preferences first
    Call ExportPreferencesConfig

    ' Step 4: Create backup of current workbook
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim currentPath As String
    Dim backupPath As String
    Dim timestamp As String
    currentPath = ThisWorkbook.FullName
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    backupPath = Replace(currentPath, ".xlsm", "_backup_" & timestamp & ".xlsm")

    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupPath
    Application.DisplayAlerts = True

    ' Step 5: Run Python rebuild
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim buildCmd As String
    Dim yearVal As String
    yearVal = CStr(GetReportYear())
    buildCmd = "cmd /c cd /d """ & ThisWorkbook.Path & """ && python build_workbook.py --year " & yearVal

    ' Create a log file for the build process
    Dim logFile As String
    logFile = ThisWorkbook.Path & "\\rebuild_log.txt"
    buildCmd = buildCmd & " > """ & logFile & """ 2>&1"

    ' Show progress message
    Application.StatusBar = "Rebuilding workbook... Please wait..."

    ' Run the command and wait
    Dim exitCode As Long
    exitCode = shell.Run(buildCmd, 0, True)

    Application.StatusBar = False

    ' Step 6: Check results
    If exitCode = 0 Then
        ' Success
        MsgBox "Rebuild successful!" & vbCrLf & vbCrLf & _
               "Backup saved to:" & vbCrLf & backupPath & vbCrLf & vbCrLf & _
               "NEXT STEPS:" & vbCrLf & _
               "1. This workbook will close" & vbCrLf & _
               "2. Open the new Bed_Utilization_" & yearVal & ".xlsm" & vbCrLf & _
               "3. Click 'IMPORT OLD WORKBOOK' button" & vbCrLf & _
               "4. Select the backup file to restore your data", _
               vbInformation, "Rebuild Complete"

        ' Close current workbook without saving (backup already created)
        ThisWorkbook.Close SaveChanges:=False
    Else
        ' Failed - read log file
        Dim errorMsg As String
        errorMsg = "Rebuild FAILED (Exit code: " & exitCode & ")" & vbCrLf & vbCrLf

        If fso.FileExists(logFile) Then
            Dim ts As Object
            Set ts = fso.OpenTextFile(logFile, 1)
            Dim logContent As String
            logContent = ts.ReadAll
            ts.Close

            ' Show last 500 characters of log
            If Len(logContent) > 500 Then
                logContent = "..." & Right(logContent, 500)
            End If
            errorMsg = errorMsg & "Build log:" & vbCrLf & logContent
        Else
            errorMsg = errorMsg & "No log file generated."
        End If

        MsgBox errorMsg, vbCritical, "Rebuild Failed"

        ' Keep current workbook open
        ' User can try again or manually fix issues
    End If

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    MsgBox "Error during rebuild: " & Err.Description & vbCrLf & vbCrLf & _
           "Your current workbook was not modified." & vbCrLf & _
           "Backup file: " & backupPath, _
           vbCritical, "Rebuild Error"
End Sub
