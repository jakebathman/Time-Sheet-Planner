VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarFrm 
   Caption         =   "Calendar Control"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3960
   OleObjectBlob   =   "CalendarFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CalendarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Dim ThisDay As Date
Dim ThisYear, ThisMth As Date
Dim CreateCal As Boolean
Dim i As Integer





Private Sub btnTodaySelect_Click()
    ActiveCell.Value = Date
    Me.Hide
End Sub

Private Sub btnTomorrowSelect_Click()
    ActiveCell.Value = Date + 1
    Me.Hide
End Sub






Private Sub UserForm_Activate()

Me.Top = ActiveCell.Top + (Application.Height - Application.UsableHeight) + ActiveCell.Height + Application.Top

Me.Left = Application.Left + ActiveCell.Left + (ActiveWindow.Width - ActiveWindow.UsableWidth)


End Sub

Private Sub UserForm_Initialize()
    Application.EnableEvents = False
    'starts the form on todays date
    ThisDay = Date
    ThisMth = Format(ThisDay, "mm")
    ThisYear = Format(ThisDay, "yyyy")
    For i = 1 To 12
        CB_Mth.AddItem Format(DateSerial(Year(Date), Month(Date) + i, 0), "mmmm")
    Next
    CB_Mth.ListIndex = Format(Date, "mm") - Format(Date, "mm")
    For i = -20 To 50
        If i = 1 Then CB_Yr.AddItem Format((ThisDay), "yyyy") Else CB_Yr.AddItem _
           Format((DateAdd("yyyy", (i - 1), ThisDay)), "yyyy")
    Next
    CB_Yr.ListIndex = 21
    'Builds the calendar with todays date
    CalendarFrm.Width = CalendarFrm.Width
    CreateCal = True
    Call Build_Calendar
    Application.EnableEvents = True
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

End Sub
Private Sub CB_Mth_Change()
    'rebuilds the calendar when the month is changed by the user
    Build_Calendar
End Sub
Private Sub CB_Yr_Change()
    'rebuilds the calendar when the year is changed by the user
    Build_Calendar
End Sub
Private Sub Build_Calendar()
    'the routine that actually builds the calendar each time
    If CreateCal = True Then
        CalendarFrm.Caption = " " & CB_Mth.Value & " " & CB_Yr.Value
        'sets the focus for the todays date button
        CommandButton1.SetFocus
        For i = 1 To 42
            If i < Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value)) Then
                Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                                                             ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "d")
                Controls("D" & (i)).ControlTipText = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                                                                    ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "m/d/yy")
            ElseIf i >= Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value)) Then
                Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) _
                                                                             & "/1/" & (CB_Yr.Value))), ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "d")
                Controls("D" & (i)).ControlTipText = Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                                                                    ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "m/d/yy")
            End If
            If Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                              ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "mmmm") = ((CB_Mth.Value)) Then
                If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H80000018  '&H80000010
                Controls("D" & (i)).Font.Bold = True
                If Format(DateAdd("d", (i - Weekday((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), _
                                  ((CB_Mth.Value) & "/1/" & (CB_Yr.Value))), "m/d/yy") = Format(ThisDay, "m/d/yy") Then Controls("D" & (i)).SetFocus
            Else
                If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H8000000F
                Controls("D" & (i)).Font.Bold = False
            End If
        Next
    End If
End Sub
Private Sub D1_Click()
    ActiveCell.Value = D1.ControlTipText
    Me.Hide

End Sub
Private Sub D2_Click()
    ActiveCell.Value = D2.ControlTipText
    Me.Hide

End Sub
Private Sub D3_Click()
    ActiveCell.Value = D3.ControlTipText
    Me.Hide

End Sub
Private Sub D4_Click()
    ActiveCell.Value = D4.ControlTipText
    Me.Hide

End Sub
Private Sub D5_Click()
    ActiveCell.Value = D5.ControlTipText
    Me.Hide

End Sub
Private Sub D6_Click()
    ActiveCell.Value = D6.ControlTipText
    Me.Hide

End Sub
Private Sub D7_Click()
    ActiveCell.Value = D7.ControlTipText
    Me.Hide

End Sub
Private Sub D8_Click()
    ActiveCell.Value = D8.ControlTipText
    Me.Hide

End Sub
Private Sub D9_Click()
    ActiveCell.Value = D9.ControlTipText
    Me.Hide

End Sub
Private Sub D10_Click()
    ActiveCell.Value = D10.ControlTipText
    Me.Hide

End Sub
Private Sub D11_Click()
    ActiveCell.Value = D11.ControlTipText
    Me.Hide

End Sub
Private Sub D12_Click()
    ActiveCell.Value = D12.ControlTipText
    Me.Hide

End Sub
Private Sub D13_Click()
    ActiveCell.Value = D13.ControlTipText
    Me.Hide

End Sub
Private Sub D14_Click()
    ActiveCell.Value = D14.ControlTipText
    Me.Hide

End Sub
Private Sub D15_Click()
    ActiveCell.Value = D15.ControlTipText
    Me.Hide

End Sub
Private Sub D16_Click()
    ActiveCell.Value = D16.ControlTipText
    Me.Hide

End Sub
Private Sub D17_Click()
    ActiveCell.Value = D17.ControlTipText
    Me.Hide

End Sub
Private Sub D18_Click()
    ActiveCell.Value = D18.ControlTipText
    Me.Hide

End Sub
Private Sub D19_Click()
    ActiveCell.Value = D19.ControlTipText
    Me.Hide

End Sub
Private Sub D20_Click()
    ActiveCell.Value = D20.ControlTipText
    Me.Hide

End Sub
Private Sub D21_Click()
    ActiveCell.Value = D21.ControlTipText
    Me.Hide

End Sub
Private Sub D22_Click()
    ActiveCell.Value = D22.ControlTipText
    Me.Hide

End Sub
Private Sub D23_Click()
    ActiveCell.Value = D23.ControlTipText
    Me.Hide

End Sub
Private Sub D24_Click()
    ActiveCell.Value = D24.ControlTipText
    Me.Hide

End Sub
Private Sub D25_Click()
    ActiveCell.Value = D25.ControlTipText
    Me.Hide

End Sub
Private Sub D26_Click()
    ActiveCell.Value = D26.ControlTipText
    Me.Hide

End Sub
Private Sub D27_Click()
    ActiveCell.Value = D27.ControlTipText
    Me.Hide

End Sub
Private Sub D28_Click()
    ActiveCell.Value = D28.ControlTipText
    Me.Hide

End Sub
Private Sub D29_Click()
    ActiveCell.Value = D29.ControlTipText
    Me.Hide

End Sub
Private Sub D30_Click()
    ActiveCell.Value = D30.ControlTipText
    Me.Hide

End Sub
Private Sub D31_Click()
    ActiveCell.Value = D31.ControlTipText
    Me.Hide

End Sub
Private Sub D32_Click()
    ActiveCell.Value = D32.ControlTipText
    Me.Hide

End Sub
Private Sub D33_Click()
    ActiveCell.Value = D33.ControlTipText
    Me.Hide

End Sub
Private Sub D34_Click()
    ActiveCell.Value = D34.ControlTipText
    Me.Hide

End Sub
Private Sub D35_Click()
    ActiveCell.Value = D35.ControlTipText
    Me.Hide

End Sub
Private Sub D36_Click()
    ActiveCell.Value = D36.ControlTipText
    Me.Hide

End Sub
Private Sub D37_Click()
    ActiveCell.Value = D37.ControlTipText
    Me.Hide

End Sub
Private Sub D38_Click()
    ActiveCell.Value = D38.ControlTipText
    Me.Hide

End Sub
Private Sub D39_Click()
    ActiveCell.Value = D39.ControlTipText
    Me.Hide

End Sub
Private Sub D40_Click()
    ActiveCell.Value = D40.ControlTipText
    Me.Hide

End Sub
Private Sub D41_Click()
    ActiveCell.Value = D41.ControlTipText
    Me.Hide

End Sub
Private Sub D42_Click()
    ActiveCell.Value = D42.ControlTipText
    Me.Hide

End Sub


