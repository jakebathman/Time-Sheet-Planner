VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub btnPrintTimeOffForm_Click()
    'ActiveSheet.PrintPreview
    Sheets("Time Off Form").PrintOut From:=1, To:=1
End Sub

Private Sub calToDate_Click()
    Range("$C$6").Value = calToDate.Value
End Sub

Private Sub calFromDate_Click()
    Range("$C$4").Value = calFromDate.Value
End Sub

Private Sub chkCat_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkCat.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkCat = boolThisCheck
End Sub

Private Sub chkComp_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkComp.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkComp = boolThisCheck
End Sub

Private Sub chkCompEarned_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkCompEarned.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkCompEarned = boolThisCheck
End Sub

Private Sub chkOther_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkOther.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkOther = boolThisCheck
End Sub

Private Sub chkOvertimeEarned_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkOvertimeEarned.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkOvertimeEarned = boolThisCheck
End Sub

Private Sub chkPTO_Click()
    Dim boolThisCheck As Boolean
    
    boolThisCheck = Me.chkPTO.Value

    chkPTO.Value = False
    chkCat.Value = False
    chkComp.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    
    chkPTO = boolThisCheck
End Sub

Private Sub Worksheet_Activate()
    Dim i%
    Dim v

    ' My code
    With Sheets("References")
        For i = 2 To 50
            If .Cells(i, 1).Value = vbNullString Then
                Sheets("Time Off Form").boxEmployeeName.ListFillRange = "References!$A$2:$A$" & i - 1
                Exit For
            End If
        Next i
    End With

End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If ActiveWindow.View <> xlNormalView Then ActiveWindow.View = xlNormalView
    calDateSubmitted.Visible = False
    calFromDate.Visible = False
    calToDate.Visible = False

    If Target.Address = "$C$4" Then
        If Target.Value = vbNullString Then
            calFromDate.Value = Date
            Target.Value = Date
        Else
            calFromDate.Value = Target.Value
        End If
        calFromDate.Visible = True
    ElseIf Target.Address = "$C$6" Then
        If Target.Value = vbNullString Then
            calToDate.Value = Date
            Target.Value = Date
        Else
            calToDate.Value = Target.Value
        End If
        calToDate.Visible = True
    ElseIf Target.Address = "$H$2" Then
        If Target.Value = vbNullString Then
            calDateSubmitted.Value = Date
            Target.Value = Date
        Else
            calDateSubmitted.Value = Target.Value
        End If
        calDateSubmitted.Visible = True
    End If

End Sub




Private Sub Worksheet_Change(ByVal Target As Range)
    If Range("$C$6").Value < Range("$C$4").Value Then Range("$C$6").Value = Range("$C$4").Value

    ' Code to allow flexible time input formats
    Dim dblTimeInput#
    Dim strTimeInput$
    Dim tTime As Date
    Dim i%, intLeftCol%, intRightCol%, c%
    Dim strLeftCol$, strRightCol$
    Dim dblTimeOutput#
    Dim vNew
    Dim vOld
    Dim intLenStrTimeInput%
    Dim rngWorkingCell As Range
    Dim boolNumeric As Boolean

    Application.ScreenUpdating = False
    If Target.Count > 1 Then GoTo EndOfSub

    For i = 1 To 50
        If Cells(1, i).Value = vbNullString And Cells(1, i + 1).Value = vbNullString And Cells(1, i + 2).Value = vbNullString Then Exit For
        If Cells(1, i).Value = "In 6/27" Then intLeftCol = i: strLeftCol = Mid(Cells(1, i).Address(1, 1, xlA1), 2, 1)
        If Cells(1, i).Value = "Out 6/27" Then intRightCol = i: strRightCol = Mid(Cells(1, i).Address(1, 1, xlA1), 2, 1): Exit For
    Next i
    If Target.Address = "$D$4" Or Target.Address = "$D$6" Then   'checks that the edited cell is in the time punch range

        On Error Resume Next
        If Target.Value = "" Or Target.Value = " " Or Target.Value = 0 Or Target.Value = Empty Then
            Application.EnableEvents = False
            Target.Value = ""
            Application.EnableEvents = True
            GoTo EndOfSub    'kills if the user just deleted or blanked a cell within the range
        End If

        strTimeInput = Target.Value
        ' Change input back to string, to manipulate it a bit
        If CDbl(strTimeInput) < 1 Then
            strTimeInput = CStr(Floor(Target.Value * 24))
            strTimeInput = strTimeInput & Floor(((Target.Value * 24) - CDbl(strTimeInput)) * 60)
        End If
        dblTimeInput = CDbl(strTimeInput) / 24 / 100
        intLenStrTimeInput = Len(strTimeInput)
        On Error GoTo 0

        If InStr(1, strTimeInput, "p", vbTextCompare) > 0 Then
            strTimeInput = CStr((CInt(Left(strTimeInput, 1)) + 12)) & Mid(strTimeInput, 2)
            strTimeInput = Left(strTimeInput, InStr(1, strTimeInput, "p", vbTextCompare) - 1)
        End If
        If InStr(1, strTimeInput, "a", vbTextCompare) > 0 Then
            strTimeInput = Left(strTimeInput, InStr(1, strTimeInput, "a", vbTextCompare) - 1)
        End If
        boolNumeric = False
        While Not boolNumeric
            For c = 1 To Len(strTimeInput)
                If Not IsNumeric(Mid(strTimeInput, c, 1)) Then
                    strTimeInput = Left(strTimeInput, c - 1) & Mid(strTimeInput, c + 1)
                    Exit For
                End If
            Next c
            If c = Len(strTimeInput) + 1 Then boolNumeric = True Else boolNumeric = False
        Wend

        If Left(strTimeInput, 1) = 0 And InStr(1, strTimeInput, ".", vbTextCompare) = 0 Then
            strTimeInput = Trim(Mid(strTimeInput, 2))
            dblTimeInput = CDbl(strTimeInput)
        End If

        If strTimeInput = "" Or (InStr(1, strTimeInput, ".", vbTextCompare) > 0 And dblTimeInput < 1) Then
            Application.EnableEvents = False
            If strTimeInput = "" Then Target.Value = ""
            Application.EnableEvents = True
            GoTo EndOfSub    'kills if the user just deleted or blanked a cell within the range
        End If


        Dim intAdjustor As Integer
        Dim strH$, strM$, strS$
        intAdjustor = 0
        Application.EnableEvents = False
        Range(Target.Address) = ""
        Application.EnableEvents = True
        Select Case Len(strTimeInput)
            Case Is = 6
                strH = Left(strTimeInput, 2)
                strM = Mid(strTimeInput, 3, 2)
                strS = Right(strTimeInput, 2)
                dblTimeOutput = ((CDbl(strH) + intAdjustor) / 24) + (((CDbl(strM) * 100) / 60) / 2400) + ((((CDbl(strS) * 100) / 60) / 60) / 2400)
            Case Is = 5
                strH = Left(strTimeInput, 1)
                strM = Mid(strTimeInput, 1, 2)
                strS = Right(strTimeInput, 2)
                dblTimeOutput = ((CDbl(strH) + intAdjustor) / 24) + (((CDbl(strM) * 100) / 60) / 2400) + ((((CDbl(strS) * 100) / 60) / 60) / 2400)
            Case Is = 4
                strH = Left(strTimeInput, 2)
                strM = Right(strTimeInput, 2)
                strS = "00"
                If (CDbl(strM) = 7) Or (CDbl(strM) = 22) Or (CDbl(strM) = 37) Or (CDbl(strM) = 52) Then
                    strS = InputBox("This may round differently with seconds included. How many seconds? (leave blank if :00)", "Seconds?")
                    If strS = vbNullString Or Not IsNumeric(strS) Then strS = "00"
                End If
                dblTimeOutput = ((CDbl(strH) + intAdjustor) / 24) + (((CDbl(strM) * 100) / 60) / 2400) + ((((CDbl(strS) * 100) / 60) / 60) / 2400)
            Case Is = 3
                strH = Left(strTimeInput, 1)
                strM = Right(strTimeInput, 2)
                strS = "00"
                If (CDbl(strM) = 7) Or (CDbl(strM) = 22) Or (CDbl(strM) = 37) Or (CDbl(strM) = 52) Then
                    strS = InputBox("This may round differently with seconds included. How many seconds? (leave blank if :00)", "Seconds?")
                    If strS = vbNullString Or Not IsNumeric(strS) Then strS = "00"
                End If
                dblTimeOutput = ((CDbl(strH) + intAdjustor) / 24) + (((CDbl(strM) * 100) / 60) / 2400) + ((((CDbl(strS) * 100) / 60) / 60) / 2400)
            Case Is = 2
                If CDbl(strTimeInput) > 24 Then Target.Value = "": GoTo EndOfSub Else dblTimeOutput = ((CDbl(strTimeInput) + intAdjustor) / 24)
            Case Is = 1
                dblTimeOutput = ((CDbl(strTimeInput) + intAdjustor) / 24)
        End Select
        Application.EnableEvents = False
        Range(Target.Address).Value = dblTimeOutput
        Range(Target.Address).NumberFormat = "h:mm AM/PM;@"
    End If

EndOfSub:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub


End Sub
