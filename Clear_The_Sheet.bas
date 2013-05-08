Attribute VB_Name = "Clear_The_Sheet"
Option Explicit
Public YesNo As Variant

Sub clearstuff()

    Dim i As Integer
    Dim t
    Dim x#

    With frmWorking
        .Show False
        .Height = 60
        .Width = 245
        .Top = Application.Top + (Application.Height / 2) - (.Height / 2) - 75
        .Left = Application.Left + (Application.Width / 2) - (.Width / 2)
        .Label2.Caption = vbNullString
    End With

    Call UpdateWorkingForm

    Call MaintenanceAndRepair

    Call FindTheSheetInfo

    YesNo = MsgBox("Really clear your inputted time below?", vbOKCancel)

    If YesNo <> vbOK Then Unload frmWorking: Exit Sub
    Application.ScreenUpdating = False

    Dim boolOverwriteBackup
    Dim boolBackupOperationComplete As Boolean
    Dim vbRUSure

    Dim j As Integer
    Dim intCurSheetNum As Integer

    dblTimeNowRnd = Now()


    'create backups
    boolBackupOperationComplete = False
    Application.DisplayAlerts = False

    On Error Resume Next
    IsError (Sheets("Backup of Time Sheet Planner").Index)

    If Err.Number <> 9 Then
        Call UpdateWorkingForm
        While boolBackupOperationComplete = False
            boolOverwriteBackup = MsgBox("Backup of main sheet already exists. Overwrite?", vbYesNoCancel)
            Select Case boolOverwriteBackup
                Case vbYes
                    'creates hidden backup sheet of backup, just in case. Not accessed anywhere else, must be manually reinstated
                    On Error Resume Next
                    Sheets("Hidden Backup of Old Backup").Visible = True
                    Sheets("Hidden Backup of Old Backup").Delete
                    On Error GoTo ErrHandlerCode
                    Sheets("Backup of Time Sheet Planner").Activate
                    Sheets("Backup of Time Sheet Planner").Copy after:=Sheets("Backup of Time Sheet Planner")
                    ActiveSheet.Name = "TmpSheet" & dblTimeNowRnd + Rnd
                    intCurSheetNum = ActiveSheet.Index
                    ActiveSheet.Name = "Hidden Backup of Old Backup"
                    ActiveSheet.Visible = xlSheetVeryHidden
                    Sheets("Backup of Time Sheet Planner").Delete
                    Sheets(1).Activate
                    Sheets(1).Copy after:=Sheets(1)
                    ActiveSheet.Name = "Backup of Time Sheet Planner"
                    boolBackupOperationComplete = True
                Case vbNo
                    vbRUSure = MsgBox("This won't back up your current sheet, but keep a (possibly) old backup. Really continue without backing up main sheet?", vbYesNo + vbQuestion)
                    If vbRUSure = vbYes Then boolBackupOperationComplete = True
                Case vbCancel
                    End
            End Select
        Wend
    Else
        Call UpdateWorkingForm
        Sheets("Time Sheet Planner").Activate
        ActiveSheet.Copy after:=Sheets("Time Sheet Planner")
        ActiveSheet.Name = "Backup of Time Sheet Planner"
        Call UpdateWorkingForm
    End If

    'creates hidden backup of main sheet, just in case. Not accessed anywhere else, must be manually reinstated
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Hidden Backup of Main").Visible = True
    Sheets("Hidden Backup of Main").Delete
    On Error GoTo ErrHandlerCode
    Call UpdateWorkingForm
    Sheets("Time Sheet Planner").Activate
    Sheets("Time Sheet Planner").Copy after:=Sheets("Backup of Time Sheet Planner")
    Call UpdateWorkingForm
    ActiveSheet.Name = "TmpSheet" & dblTimeNowRnd + Rnd
    intCurSheetNum = ActiveSheet.Index
    ActiveSheet.Name = "Hidden Backup of Main"
    ActiveSheet.Visible = xlSheetVeryHidden
    Call UpdateWorkingForm

    Application.DisplayAlerts = True


    Call UpdateWorkingForm



    Application.EnableEvents = False

    Sheets("Time Sheet Planner").Activate

    If YesNo = vbOK Then
        Call UpdateWorkingForm(25)
        With Sheets("Time Sheet Planner").Range("B3:I14")
            .Value = ""
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            Call UpdateWorkingForm(30)
            x = 30
            t = Timer
            While Timer < t + 0.45
                Call UpdateWorkingForm(x)
                x = x + 0.05
                'Debug.Print x
                DoEvents
            Wend

            On Error Resume Next
            .Comment.Delete
            On Error GoTo ErrHandlerCode
        End With
        With Sheets("Time Sheet Planner").Range("K3:K14")
            .Value = ""
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            Call UpdateWorkingForm(30)
            t = Timer
            While Timer < t + 0.15
                Call UpdateWorkingForm(x)
                x = x + 0.05
                'Debug.Print x
                DoEvents
            Wend

            On Error Resume Next
            .Comment.Delete
            On Error GoTo ErrHandlerCode
        End With
        Sheets(1).Cells(23, 2).Activate
        Sheets(1).Cells(23, 2) = ""
    Else
        Call UpdateWorkingForm
        End
    End If

    Call UpdateWorkingForm(75)

    With Cells(17, 2)
        .Value = ""
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
            While Timer < t + 0.1
                Call UpdateWorkingForm(x)
                x = x + 0.04
                Debug.Print x
                DoEvents
            Wend

    Cells(3, 2).Select
    Call UpdateWorkingForm

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call UpdateWorkingForm(100)


    Dim boolRestoreBackup

    boolRestoreBackup = MsgBox("All clear, captain!" & vbCrLf & vbCrLf & "I created a backup (just in case) of your original data. " _
                             & "Delete the backup now and use this empty sheet?" _
                             & vbCrLf & vbCrLf & "Yes: Keep changes to main sheet and delete backup." _
                             & vbCrLf & "No: Restore my old stuff (Undo clearing data)." _
                             & vbCrLf & vbCrLf & "(WARNING! Any action is not reversible.)", vbYesNo)

    Select Case boolRestoreBackup
        Case vbYes
            'Delete backup sheet
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Backup of Time Sheet Planner").Delete
            Application.DisplayAlerts = True
        Case vbNo
            'Restore backup sheet (pre-macro state)
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Time Sheet Planner").Delete
            Sheets("Backup of Time Sheet Planner").Name = "Time Sheet Planner"
            Application.DisplayAlerts = True
    End Select


ErrHandlerCode:
    If Err.Number <> 0 Then
        MsgBox ("Woops! I've encountered an error I didn't plan for." & vbCrLf & "Please report this error to the developer:" _
              & vbCrLf & vbCrLf & "Error # " & Str(Err.Number) & ": " & Err.Description)
    End If
    Unload frmWorking

    Call MaintenanceAndRepair


End Sub



Public Function UpdateWorkingForm(Optional dblPctTitle#)
Dim arrRotatingChar(1 To 4)
Dim iBarLeft#, iBarWidth#, iBarRight#
Dim iBGLeft#, iBGWidth#, iBGRight#
Dim iBarTwoWidth#, iBarTwoRight#
Dim NewBarRight#
Dim steps#

arrRotatingChar(1) = "|"
arrRotatingChar(2) = " | "
arrRotatingChar(3) = "/"
arrRotatingChar(4) = " / "

Select Case frmWorking.lblProgressText.Caption
    Case "|"
        frmWorking.lblProgressText.Caption = arrRotatingChar(2)
    Case "/"
        frmWorking.lblProgressText.Caption = arrRotatingChar(3)
    Case "--"
        frmWorking.lblProgressText.Caption = arrRotatingChar(4)
    Case "\"
        frmWorking.lblProgressText.Caption = arrRotatingChar(1)
End Select

iBarLeft = frmWorking.lblMovingBar.Left
iBarWidth = frmWorking.lblMovingBar.Width
iBarRight = iBarLeft + iBarWidth

iBGWidth = frmWorking.Label3.Width
iBGLeft = frmWorking.Label3.Left
iBGRight = iBGWidth + iBGLeft

iBarTwoWidth = frmWorking.lblMoving2.Width
iBarTwoRight = iBarTwoWidth + 10

steps = Round((iBGWidth / 47), 0)

If Round(iBarRight + steps + 1, 0) > iBGRight Then
    If Round(iBarLeft + steps + 1, 0) > iBGRight Then 'reset bar to the left
        If iBarTwoWidth > 0 Then
            frmWorking.lblMoving2.Width = 0
            frmWorking.lblMovingBar.Left = steps + 10
            frmWorking.lblMovingBar.Width = 85
        Else
            frmWorking.lblMovingBar.Left = 10
            frmWorking.lblMovingBar.Width = 85
            frmWorking.lblMoving2.Width = 0
        End If
    Else
        frmWorking.lblMovingBar.Left = iBarLeft + steps
        frmWorking.lblMovingBar.Width = iBGRight - (iBarLeft + steps) - 2
        NewBarRight = frmWorking.lblMovingBar.Left + 85 'measures new width of green bar if spills over
        frmWorking.lblMoving2.Width = (NewBarRight - iBGRight)
    End If
Else
    frmWorking.lblMovingBar.Left = iBarLeft + steps
End If

If dblPctTitle > 0 Then frmWorking.Caption = Round(dblPctTitle, 1) & "% Complete"
If dblPctTitle = 200 Then
    frmWorking.Caption = "100% Done!"
    frmWorking.lblMovingBar.Left = 10
    frmWorking.lblMovingBar.Width = frmWorking.Label3.Width - 4
End If

frmWorking.Repaint

End Function




