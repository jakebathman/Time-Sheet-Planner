Attribute VB_Name = "Clear_The_Sheet"
Option Explicit
Public YesNo As Variant

Sub clearstuff()

    Dim i As Integer
    Dim t
    Dim X#
    Dim intNoPromptsPrefRow%
    Dim boolNoPrompts As Boolean

    boolNoPrompts = False

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

    With Sheets("User Preferences")
        For i = 1 To 50
            If InStr(1, .Cells(i, 1).Value, "No Prompts", vbTextCompare) > 0 Then
                intNoPromptsPrefRow = i
                Exit For
            End If
        Next i
        If StrComp(.Cells(i, 3).Value, "X", vbTextCompare) = 0 And .Cells(i + 1, 3).Value = vbNullString Then
            boolNoPrompts = True
        End If
    End With

    If boolNoPrompts Then
        YesNo = vbOK
    Else
        YesNo = MsgBox("Really clear your inputted time below?", vbOKCancel)
    End If

    If YesNo <> vbOK Then Unload frmWorking: Exit Sub
    Application.ScreenUpdating = False

    Dim boolOverwriteBackup
    Dim boolBackupOperationComplete As Boolean
    Dim vbRUSure

    Dim j As Integer
    Dim intCurSheetNum As Integer

    dblTimeNowRnd = Now()

    Application.DisplayAlerts = True

    Call UpdateWorkingForm

    Application.EnableEvents = False

    Sheets("Time Sheet Planner").Activate

    If YesNo = vbOK Then
        Call UpdateWorkingForm(25)
        With Sheets("Time Sheet Planner").Range("B3:I9")
            .Value = ""
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            Call UpdateWorkingForm(30)
            X = 30
            t = Timer
            While Timer < t + 0.1
                Call UpdateWorkingForm(X)
                X = X + 0.05
                'Debug.Print x
                DoEvents
            Wend

            On Error Resume Next
            .Comment.Delete
            On Error GoTo ErrHandlerCode
        End With
        With Sheets("Time Sheet Planner").Range("L3:L9")
            .Value = ""
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            Call UpdateWorkingForm(30)
            t = Timer
            While Timer < t + 0.1
                Call UpdateWorkingForm(X)
                X = X + 0.05
                'Debug.Print x
                DoEvents
            Wend

            On Error Resume Next
            .Comment.Delete
            On Error GoTo ErrHandlerCode
        End With
    Else
        Call UpdateWorkingForm
        End
    End If

    Call UpdateWorkingForm(75)

    ' Add drop-down for time off code selection
    With Range("H3:H9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, Formula1:="=References!$B$2:$B$5"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = False
        .ShowError = False
    End With


    '    With Cells(17, 2)
    '        .Value = ""
    '        .Interior.Pattern = xlNone
    '        .Interior.TintAndShade = 0
    '        .Interior.PatternTintAndShade = 0
    '    End With
    While Timer < t + 0.1
        Call UpdateWorkingForm(X)
        X = X + 0.04
        Debug.Print X
        DoEvents
    Wend

    Sheets("Time Sheet Planner").btnCreateTimeOffSheet.Visible = False
    Sheets("Time Sheet Planner").btnCreateCompForm.Visible = False

    Cells(3, 2).Select
    Call UpdateWorkingForm

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call UpdateWorkingForm(100)


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
        If Round(iBarLeft + steps + 1, 0) > iBGRight Then    'reset bar to the left
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
            NewBarRight = frmWorking.lblMovingBar.Left + 85    'measures new width of green bar if spills over
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




