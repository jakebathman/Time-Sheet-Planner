VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFlagPunchesForEmail 
   Caption         =   "Highlight these punches"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   OleObjectBlob   =   "frmFlagPunchesForEmail.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmFlagPunchesForEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer
Public boolTFSwitch As Boolean
Public boolEnabledSwitch As Boolean

Private Sub btnCancel_Click()
    End
End Sub

Private Sub btnSubmit_Click()
    Dim j As Integer
    Dim o As Integer
    Dim k As Integer

    On Error Resume Next
    For i = 1 To 48
        j = 1
        k = 1
        Select Case i
            Case 7, 14, 21, 28, 35, 42
                'foo
            Case 1 To 6
                o = 0
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Monday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j
            Case 8 To 13
                o = 7
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Tuesday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

            Case 15 To 20
                o = 14
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Wednesday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

            Case 22 To 27
                o = 21
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Thursday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

            Case 29 To 34
                o = 28
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Friday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

            Case 36 To 41
                o = 35
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Saturday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

            Case 43 To 48
                o = 42
                For j = 1 To 6
                    If frmFlagPunchesForEmail.Controls("CheckBox" & j + o).Value = True Then
                        k = WorksheetFunction.Match("Sunday", Range("A1:A8"), 0)
                        Cells(k, j + 1).Font.Color = RGB(250, 0, 0)
                    End If
                Next j

        End Select
    Next i

    On Error GoTo 0

    boolRedPunches = True

    Me.Hide
End Sub

Private Sub chkFriAll_Click()
    If Me.chkFriAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 29 To 34
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkMonAll_Click()
    If Me.chkMonAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 1 To 6
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkSatAll_Click()
    If Me.chkSatAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 36 To 41
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkSunAll_Click()
    If Me.chkSunAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 43 To 48
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkThuAll_Click()
    If Me.chkThuAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 22 To 27
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkTueAll_Click()
    If Me.chkTueAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 8 To 13
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub

Private Sub chkWedAll_Click()
    If Me.chkWedAll.Value = False Then boolTFSwitch = False Else boolTFSwitch = True
    For i = 15 To 20
        With Me.Controls("CheckBox" & i)
            If .Enabled = False Then
                'foo
            Else
                .Locked = boolTFSwitch
                .Value = boolTFSwitch
            End If
        End With
    Next i
End Sub




Sub UserForm_Activate()
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

    Dim intDayOffset As Integer
    Dim j As Integer
    Dim boolMon As Boolean
    Dim boolTue As Boolean
    Dim boolWed As Boolean
    Dim boolThu As Boolean
    Dim boolFri As Boolean
    Dim boolSat As Boolean
    Dim boolSun As Boolean

    boolMon = False
    boolTue = False
    boolWed = False
    boolThu = False
    boolFri = False
    boolSat = False
    boolSun = False

    For i = 1 To 48
        Select Case i
            Case 7, 14, 21, 28, 35, 42

                'foo
            Case Else
                frmFlagPunchesForEmail.Controls("CheckBox" & i).Caption = ""
                frmFlagPunchesForEmail.Controls("CheckBox" & i).Enabled = False
                frmFlagPunchesForEmail.Controls("CheckBox" & i).Visible = False
                frmFlagPunchesForEmail.Controls("CheckBox" & i).Locked = False
        End Select
        frmFlagPunchesForEmail.Repaint

    Next i

    For i = 1 To WorksheetFunction.CountA(Range("A1:A10"))
        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Monday" Then
            boolMon = True
            intDayOffset = 0
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Tuesday" Then
            boolTue = True
            intDayOffset = 7
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Wednesday" Then
            boolWed = True
            intDayOffset = 14
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Thursday" Then
            boolThu = True
            intDayOffset = 21
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Friday" Then
            boolFri = True
            intDayOffset = 28
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Saturday" Then
            boolSat = True
            intDayOffset = 35
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

        If Sheets(strSheetName).Cells(i + 1, 1).Value = "Sunday" Then
            boolSun = True
            intDayOffset = 42
            For j = 1 To 6
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = MakeTimeString(Sheets(strSheetName).Cells(i + 1, j + 1))
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Visible = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = True
                frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Locked = False
                If frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "No Punch" Then frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Caption = "": frmFlagPunchesForEmail.Controls("CheckBox" & j + intDayOffset).Enabled = False
            Next j
        End If

    Next i

    With frmFlagPunchesForEmail
        If boolMon = False Then .chkMonAll.Enabled = False: .chkMonAll.Visible = False Else .chkMonAll.Enabled = True: .chkMonAll.Visible = True
        If boolTue = False Then .chkTueAll.Enabled = False: .chkTueAll.Visible = False Else .chkTueAll.Enabled = True: .chkTueAll.Visible = True
        If boolWed = False Then .chkWedAll.Enabled = False: .chkWedAll.Visible = False Else .chkWedAll.Enabled = True: .chkWedAll.Visible = True
        If boolThu = False Then .chkThuAll.Enabled = False: .chkThuAll.Visible = False Else .chkThuAll.Enabled = True: .chkThuAll.Visible = True
        If boolFri = False Then .chkFriAll.Enabled = False: .chkFriAll.Visible = False Else .chkFriAll.Enabled = True: .chkFriAll.Visible = True
        If boolSat = False Then .chkSatAll.Enabled = False: .chkSatAll.Visible = False Else .chkSatAll.Enabled = True: .chkSatAll.Visible = True
        If boolSun = False Then .chkSunAll.Enabled = False: .chkSunAll.Visible = False Else .chkSunAll.Enabled = True: .chkSunAll.Visible = True
    End With

End Sub



