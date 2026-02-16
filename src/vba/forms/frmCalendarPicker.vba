VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendarPicker
   Caption         =   "Select Date"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   OleObjectBlob   =   "frmCalendarPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendarPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' UserForm: frmCalendarPicker
' Purpose:  Visual calendar picker for date selection (64-bit compatible)
'           Replacement for MSComCtl2.DTPicker which doesn't work on 64-bit Excel
'
' Created: 2026-02-16
'
' Public Interface:
'   Function ShowCalendar(Optional initialDate) As Variant
'     - Displays calendar and returns selected date or Empty if cancelled
'
' Layout:
'   ┌──────────────────────────────────────┐
'   │  [< Prev]  February 2026  [Next >]  │
'   │  Month [▼] Year [▼]      [Today]    │
'   ├──────────────────────────────────────┤
'   │  Su  Mo  Tu  We  Th  Fr  Sa         │
'   │                          1           │
'   │   2   3   4   5   6   7   8         │
'   │   9  10  11  12  13  14  15         │
'   │  16  17  18  19  20  21  22         │
'   │  23  24  25  26  27  28             │
'   ├──────────────────────────────────────┤
'   │         [Select]    [Cancel]         │
'   └──────────────────────────────────────┘
'==============================================================================

Option Explicit

' Module-level variables
Private m_SelectedDate As Date
Private m_CurrentMonth As Integer
Private m_CurrentYear As Integer
Private m_DateSelected As Boolean
Private m_InitialDate As Date

'==============================================================================
' Public Function: ShowCalendar
' Purpose: Display the calendar picker and return selected date
'
' Parameters:
'   initialDate - Optional starting date to display
'
' Returns:
'   Variant - Selected date if user clicked Select, Empty if cancelled
'==============================================================================
Public Function ShowCalendar(Optional initialDate As Variant) As Variant
    ' Set initial date
    If Not IsMissing(initialDate) And IsDate(initialDate) Then
        m_InitialDate = CDate(initialDate)
    Else
        m_InitialDate = Date
    End If

    m_SelectedDate = m_InitialDate
    m_CurrentMonth = Month(m_InitialDate)
    m_CurrentYear = Year(m_InitialDate)
    m_DateSelected = False

    ' Initialize the calendar
    InitializeCalendar

    ' Show the form modally
    Me.Show vbModal

    ' Return result
    If m_DateSelected Then
        ShowCalendar = m_SelectedDate
    Else
        ShowCalendar = Empty
    End If
End Function

'==============================================================================
' Private Sub: InitializeCalendar
' Purpose: Set up month/year controls and populate calendar grid
'==============================================================================
Private Sub InitializeCalendar()
    ' Populate month combo
    cmbMonth.Clear
    Dim i As Integer
    For i = 1 To 12
        cmbMonth.AddItem MonthName(i)
    Next i
    cmbMonth.ListIndex = m_CurrentMonth - 1

    ' Populate year combo (2020-2030)
    cmbYear.Clear
    For i = 2020 To 2030
        cmbYear.AddItem CStr(i)
    Next i
    cmbYear.Value = CStr(m_CurrentYear)

    ' Update calendar display
    UpdateCalendar
End Sub

'==============================================================================
' Private Sub: UpdateCalendar
' Purpose: Refresh the calendar grid with days for current month/year
'==============================================================================
Private Sub UpdateCalendar()
    ' Update month/year label
    lblMonthYear.Caption = MonthName(m_CurrentMonth) & " " & m_CurrentYear

    ' Get first day of month and number of days
    Dim firstDay As Date
    firstDay = DateSerial(m_CurrentYear, m_CurrentMonth, 1)
    Dim firstDayOfWeek As Integer
    firstDayOfWeek = Weekday(firstDay, vbSunday) - 1  ' 0=Sunday, 6=Saturday

    Dim daysInMonth As Integer
    daysInMonth = Day(DateSerial(m_CurrentYear, m_CurrentMonth + 1, 0))

    ' Get days in previous month (for leading days)
    Dim prevMonth As Integer, prevYear As Integer
    If m_CurrentMonth = 1 Then
        prevMonth = 12
        prevYear = m_CurrentYear - 1
    Else
        prevMonth = m_CurrentMonth - 1
        prevYear = m_CurrentYear
    End If
    Dim daysInPrevMonth As Integer
    daysInPrevMonth = Day(DateSerial(prevYear, prevMonth + 1, 0))

    ' Today's date for highlighting
    Dim todayDay As Integer, todayMonth As Integer, todayYear As Integer
    todayDay = Day(Date)
    todayMonth = Month(Date)
    todayYear = Year(Date)

    ' Selected date for highlighting
    Dim selectedDay As Integer, selectedMonth As Integer, selectedYear As Integer
    selectedDay = Day(m_SelectedDate)
    selectedMonth = Month(m_SelectedDate)
    selectedYear = Year(m_SelectedDate)

    ' Clear and populate all day labels (6 rows × 7 columns = 42 labels)
    Dim row As Integer, col As Integer, labelIndex As Integer
    Dim dayNum As Integer, displayDay As Integer
    Dim lbl As MSForms.Label

    dayNum = 1 - firstDayOfWeek  ' Start with previous month's trailing days

    For row = 0 To 5
        For col = 0 To 6
            labelIndex = row * 7 + col

            ' Get reference to day label (lblDay_0_0, lblDay_0_1, etc.)
            Set lbl = Me.Controls("lblDay_" & row & "_" & col)

            If dayNum <= 0 Then
                ' Previous month's trailing days
                displayDay = daysInPrevMonth + dayNum
                lbl.Caption = CStr(displayDay)
                lbl.ForeColor = &H808080  ' Gray
                lbl.BackColor = &H8000000F  ' Default
                lbl.Tag = ""  ' Mark as not current month

            ElseIf dayNum > daysInMonth Then
                ' Next month's leading days
                displayDay = dayNum - daysInMonth
                lbl.Caption = CStr(displayDay)
                lbl.ForeColor = &H808080  ' Gray
                lbl.BackColor = &H8000000F  ' Default
                lbl.Tag = ""  ' Mark as not current month

            Else
                ' Current month's days
                lbl.Caption = CStr(dayNum)
                lbl.ForeColor = &H0  ' Black
                lbl.Tag = CStr(dayNum)  ' Store actual day number

                ' Highlight today
                If dayNum = todayDay And m_CurrentMonth = todayMonth And m_CurrentYear = todayYear Then
                    lbl.BackColor = &HFF8080  ' Light blue
                    lbl.ForeColor = &HFF0000  ' Blue text
                ' Highlight selected date
                ElseIf dayNum = selectedDay And m_CurrentMonth = selectedMonth And m_CurrentYear = selectedYear Then
                    lbl.BackColor = &H80FF80  ' Light green
                    lbl.ForeColor = &H0  ' Black text
                Else
                    lbl.BackColor = &H8000000F  ' Default
                End If
            End If

            dayNum = dayNum + 1
        Next col
    Next row
End Sub

'==============================================================================
' Event Handlers: Month/Year Selection
'==============================================================================
Private Sub cmbMonth_Change()
    If cmbMonth.ListIndex >= 0 Then
        m_CurrentMonth = cmbMonth.ListIndex + 1
        UpdateCalendar
    End If
End Sub

Private Sub cmbYear_Change()
    If IsNumeric(cmbYear.Value) Then
        m_CurrentYear = CInt(cmbYear.Value)
        UpdateCalendar
    End If
End Sub

'==============================================================================
' Event Handlers: Navigation Buttons
'==============================================================================
Private Sub btnPrev_Click()
    ' Go to previous month
    m_CurrentMonth = m_CurrentMonth - 1
    If m_CurrentMonth < 1 Then
        m_CurrentMonth = 12
        m_CurrentYear = m_CurrentYear - 1
    End If

    ' Update controls
    cmbMonth.ListIndex = m_CurrentMonth - 1
    cmbYear.Value = CStr(m_CurrentYear)
    UpdateCalendar
End Sub

Private Sub btnNext_Click()
    ' Go to next month
    m_CurrentMonth = m_CurrentMonth + 1
    If m_CurrentMonth > 12 Then
        m_CurrentMonth = 1
        m_CurrentYear = m_CurrentYear + 1
    End If

    ' Update controls
    cmbMonth.ListIndex = m_CurrentMonth - 1
    cmbYear.Value = CStr(m_CurrentYear)
    UpdateCalendar
End Sub

Private Sub btnToday_Click()
    ' Jump to today's date
    m_SelectedDate = Date
    m_CurrentMonth = Month(Date)
    m_CurrentYear = Year(Date)

    ' Update controls
    cmbMonth.ListIndex = m_CurrentMonth - 1
    cmbYear.Value = CStr(m_CurrentYear)
    UpdateCalendar
End Sub

'==============================================================================
' Event Handlers: Day Label Clicks (Generated for all 42 labels)
'==============================================================================
Private Sub HandleDayClick(lbl As MSForms.Label)
    ' Only process clicks on current month days
    If lbl.Tag <> "" Then
        Dim dayNum As Integer
        dayNum = CInt(lbl.Tag)
        m_SelectedDate = DateSerial(m_CurrentYear, m_CurrentMonth, dayNum)
        UpdateCalendar  ' Refresh to show selection
    End If
End Sub

' Day label click events for all 42 labels (6 rows × 7 columns)
Private Sub lblDay_0_0_Click()
    HandleDayClick lblDay_0_0
End Sub

Private Sub lblDay_0_1_Click()
    HandleDayClick lblDay_0_1
End Sub

Private Sub lblDay_0_2_Click()
    HandleDayClick lblDay_0_2
End Sub

Private Sub lblDay_0_3_Click()
    HandleDayClick lblDay_0_3
End Sub

Private Sub lblDay_0_4_Click()
    HandleDayClick lblDay_0_4
End Sub

Private Sub lblDay_0_5_Click()
    HandleDayClick lblDay_0_5
End Sub

Private Sub lblDay_0_6_Click()
    HandleDayClick lblDay_0_6
End Sub

Private Sub lblDay_1_0_Click()
    HandleDayClick lblDay_1_0
End Sub

Private Sub lblDay_1_1_Click()
    HandleDayClick lblDay_1_1
End Sub

Private Sub lblDay_1_2_Click()
    HandleDayClick lblDay_1_2
End Sub

Private Sub lblDay_1_3_Click()
    HandleDayClick lblDay_1_3
End Sub

Private Sub lblDay_1_4_Click()
    HandleDayClick lblDay_1_4
End Sub

Private Sub lblDay_1_5_Click()
    HandleDayClick lblDay_1_5
End Sub

Private Sub lblDay_1_6_Click()
    HandleDayClick lblDay_1_6
End Sub

Private Sub lblDay_2_0_Click()
    HandleDayClick lblDay_2_0
End Sub

Private Sub lblDay_2_1_Click()
    HandleDayClick lblDay_2_1
End Sub

Private Sub lblDay_2_2_Click()
    HandleDayClick lblDay_2_2
End Sub

Private Sub lblDay_2_3_Click()
    HandleDayClick lblDay_2_3
End Sub

Private Sub lblDay_2_4_Click()
    HandleDayClick lblDay_2_4
End Sub

Private Sub lblDay_2_5_Click()
    HandleDayClick lblDay_2_5
End Sub

Private Sub lblDay_2_6_Click()
    HandleDayClick lblDay_2_6
End Sub

Private Sub lblDay_3_0_Click()
    HandleDayClick lblDay_3_0
End Sub

Private Sub lblDay_3_1_Click()
    HandleDayClick lblDay_3_1
End Sub

Private Sub lblDay_3_2_Click()
    HandleDayClick lblDay_3_2
End Sub

Private Sub lblDay_3_3_Click()
    HandleDayClick lblDay_3_3
End Sub

Private Sub lblDay_3_4_Click()
    HandleDayClick lblDay_3_4
End Sub

Private Sub lblDay_3_5_Click()
    HandleDayClick lblDay_3_5
End Sub

Private Sub lblDay_3_6_Click()
    HandleDayClick lblDay_3_6
End Sub

Private Sub lblDay_4_0_Click()
    HandleDayClick lblDay_4_0
End Sub

Private Sub lblDay_4_1_Click()
    HandleDayClick lblDay_4_1
End Sub

Private Sub lblDay_4_2_Click()
    HandleDayClick lblDay_4_2
End Sub

Private Sub lblDay_4_3_Click()
    HandleDayClick lblDay_4_3
End Sub

Private Sub lblDay_4_4_Click()
    HandleDayClick lblDay_4_4
End Sub

Private Sub lblDay_4_5_Click()
    HandleDayClick lblDay_4_5
End Sub

Private Sub lblDay_4_6_Click()
    HandleDayClick lblDay_4_6
End Sub

Private Sub lblDay_5_0_Click()
    HandleDayClick lblDay_5_0
End Sub

Private Sub lblDay_5_1_Click()
    HandleDayClick lblDay_5_1
End Sub

Private Sub lblDay_5_2_Click()
    HandleDayClick lblDay_5_2
End Sub

Private Sub lblDay_5_3_Click()
    HandleDayClick lblDay_5_3
End Sub

Private Sub lblDay_5_4_Click()
    HandleDayClick lblDay_5_4
End Sub

Private Sub lblDay_5_5_Click()
    HandleDayClick lblDay_5_5
End Sub

Private Sub lblDay_5_6_Click()
    HandleDayClick lblDay_5_6
End Sub

'==============================================================================
' Event Handlers: Select/Cancel Buttons
'==============================================================================
Private Sub btnSelect_Click()
    m_DateSelected = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_DateSelected = False
    Me.Hide
End Sub

'==============================================================================
' Event Handler: Form Initialize
'==============================================================================
Private Sub UserForm_Initialize()
    ' Set form properties
    Me.Caption = "Select Date"

    ' Day labels will be created by Python injection
    ' Controls: cmbMonth, cmbYear, btnPrev, btnNext, btnToday
    '          lblMonthYear, lblDay_X_Y (42 labels), btnSelect, btnCancel
End Sub
