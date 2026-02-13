'###################################################################
'# MODULE: modNavigation
'# PURPOSE: Button event handlers for form navigation
'###################################################################

Option Explicit

'===================================================================
' FORM NAVIGATION HANDLERS
'===================================================================

Public Sub ShowDailyEntry()
    frmDailyEntry.Show
End Sub

Public Sub ShowAdmission()
    frmAdmission.Show
End Sub

Public Sub ShowDeath()
    frmDeath.Show
End Sub

Public Sub ShowAgesEntry()
    frmAgesEntry.Show
End Sub

Public Sub ShowRefreshReports()
    RefreshAllReports
End Sub

Public Sub ShowWardManager()
    frmWardManager.Show
End Sub

Public Sub ExportWardConfig()
    ExportWardsConfig
End Sub
