VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickDaysToEmail 
   Caption         =   "Which days?"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   OleObjectBlob   =   "frmPickDaysToEmail.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPickDaysToEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer
Public boolYesFiveDays As Boolean
Public boolYesSevenDays As Boolean


Private Sub CheckBox1_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(1) = Me.CheckBox1.Value
End Sub
Private Sub CheckBox2_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(2) = Me.CheckBox2.Value
End Sub
Private Sub CheckBox3_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(3) = Me.CheckBox3.Value
End Sub
Private Sub CheckBox4_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(4) = Me.CheckBox4.Value
End Sub
Private Sub CheckBox5_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(5) = Me.CheckBox5.Value
End Sub
Private Sub CheckBox6_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(6) = Me.CheckBox6.Value
End Sub
Private Sub CheckBox7_Click()
    If boolYesFiveDays = False And boolYesSevenDays = False Then arrCheckBoxStates(7) = Me.CheckBox7.Value
End Sub

Public Sub CheckBox8_Click()    'Mon-Fri
    If boolYesFiveDays = False And boolYesSevenDays = False Then
        boolYesFiveDays = True
        '    For i = 1 To 7 'capture current checked states
        '        arrCheckBoxStates(i) = Me.Controls("CheckBox" & i).Value
        '    Next i
        For i = 1 To 5
            Me.Controls("CheckBox" & i).Value = True
            Me.Controls("CheckBox" & i).Locked = True
            Me.Controls("CheckBox" & i).Enabled = False
        Next i
    ElseIf boolYesSevenDays = False Then
        boolYesFiveDays = False
        For i = 1 To 5    'restore previous checked states
            Me.Controls("CheckBox" & i).Locked = False
            Me.Controls("CheckBox" & i).Value = arrCheckBoxStates(i)
            Me.Controls("CheckBox" & i).Enabled = True
        Next i
    End If

End Sub

Private Sub CheckBox9_Click()    'All 7 days
    If boolYesSevenDays = False Then
        boolYesSevenDays = True
        '    For i = 1 To 7 'capture current checked states
        '        arrCheckBoxStates(i) = Me.Controls("CheckBox" & i).Value
        '    Next i
        For i = 1 To 7
            With Me
                .Controls("CheckBox" & i).Value = True
                .Controls("CheckBox" & i).Locked = True
                .Controls("CheckBox" & i).Enabled = False
            End With
        Next i
        Me.CheckBox8.Locked = True
        Me.CheckBox8.Value = True
        Me.CheckBox8.Enabled = False
    Else
        For i = 1 To 7    'restore previous checked states
            Me.Controls("CheckBox" & i).Value = arrCheckBoxStates(i)
            Me.Controls("CheckBox" & i).Enabled = True
            Me.Controls("CheckBox" & i).Locked = False
        Next i
        Me.CheckBox8.Locked = False
        Me.CheckBox8.Value = False
        Me.CheckBox8.Enabled = True
        boolYesFiveDays = False
        boolYesSevenDays = False
    End If
End Sub

Private Sub CommandButton1_Click()
    If optCurWeek.Value = True Then boolPreviousWeek = False
    Me.Hide
    Dim intC As Integer
    Dim intRow As Integer
    Dim boolBlankDays As Boolean
    Dim strMatchVal As String
    intC = 0
    ReDim arrDaysSelected(1 To 1)
    boolBlankDays = False
    For i = 1 To 7
        intRow = 0
        strMatchVal = ""
        With Me
            If .Controls("CheckBox" & i).Value = True Then
                intC = intC + 1
                If intC > 1 Then ReDim Preserve arrDaysSelected(1 To intC)
                arrDaysSelected(intC) = Trim(Left(.Controls("CheckBox" & i).Caption, InStr(.Controls("CheckBox" & i).Caption, " ") - 1))
                strMatchVal = Trim(Left(.Controls("CheckBox" & i).Caption, InStr(.Controls("CheckBox" & i).Caption, " ") - 1))
                On Error Resume Next
                intRow = WorksheetFunction.Match(strMatchVal, Range("A3:A15"), 0) + 2
                If intRow <> 0 And WorksheetFunction.CountA(Range("B" & intRow & ":G" & intRow)) = 0 Then
                    boolBlankDays = True
                End If
                On Error GoTo 0
            End If
        End With
    Next i
    If intC = 0 Then
        MsgBox ("You must pick at least ONE day! Try again.")
        Me.Show
    ElseIf boolBlankDays Then
        MsgBox ("At least one of the days you chose has no punches. Select again!")
        Me.Show
    Else
        Me.Hide
    End If
End Sub

Private Sub CommandButton2_Click()
    End
End Sub



Public Sub CommandButton3_Click()
    For i = 1 To 7
        arrCheckBoxStates(i) = False
    Next i
    For i = 1 To 9
        Me.Controls("CheckBox" & i).Value = False
        Me.Controls("CheckBox" & i).Locked = False
        Me.Controls("CheckBox" & i).Enabled = True
    Next i
End Sub

Private Sub optCurWeek_Click()
    With Me
        '.optPrevWeek.Value = False
        '.optCurWeek.Value = True
        For i = 1 To 7
            .Controls("CheckBox" & i).Caption = WeekdayName(i, False, vbMonday) & " " & Left(CStr(SetDateStrings(Date + 7) + i - 1), Len(CStr(SetDateStrings(Date + 7) + i - 1)) - 5)
        Next i
    End With

End Sub

Private Sub optPrevWeek_Click()
    With Me
        '.optPrevWeek.Value = True
        '.optCurWeek.Value = False
        For i = 1 To 7
            .Controls("CheckBox" & i).Caption = WeekdayName(i, False, vbMonday) & " " & Left(CStr(SetDateStrings(Date) + i - 1), Len(CStr(SetDateStrings(Date) + i - 1)) - 5)
        Next i
    End With


End Sub

Private Sub UserForm_Initialize()
    If boolPreviousWeek = True Then c = 7 Else c = 0
    For i = 1 To 7
        Me.Controls("CheckBox" & i).Value = False
        arrCheckBoxStates(i) = False
    Next i
    strMonth = MonthName(Month(Date))
    strDayofMonth = Day(Date)
    frmPickDaysToEmail.optCurWeek.Caption = "Current Week (Starting " & SetDateStrings(Date + 7) & ")"
    frmPickDaysToEmail.optPrevWeek.Caption = "Previous Week (Starting " & SetDateStrings(Date) & ")"
    frmPickDaysToEmail.lblTodayIs.Caption = "Today is " & WeekdayName(Weekday(Date, vbMonday), False, vbMonday) & ", " & Left(strMonth, 3) & ". " & strDayofMonth    '
    Me.optPrevWeek.Value = True
    '    Me.Show
End Sub

Private Sub UserForm_Activate()
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    ''    Me.Left = 500
    ''    Me.Top = 200
    ''    If boolPreviousWeek = True Then c = 7 Else c = 0
    ''    strMonth = MonthName(Month(Date))
    ''    strDayofMonth = Day(Date)
    ''    frmPickDaysToEmail.optCurWeek.Caption = "Current Week (Starting " & SetDateStrings(Date + 7) & ")"
    ''    frmPickDaysToEmail.optPrevWeek.Caption = "Previous Week (Starting " & SetDateStrings(Date) & ")"
    ''    frmPickDaysToEmail.lblTodayIs.Caption = "Today is " & WeekdayName(Weekday(Date, vbMonday), False, vbMonday) & ", " & Left(strMonth, 3) & ". " & strDayofMonth
End Sub
