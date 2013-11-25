VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnCreateCompForm_Click()
    
    Call CreateTimeOffForm

End Sub

Private Sub btnCreateTimeOffSheet_Click()

    Call CreateTimeOffForm

End Sub

Private Sub btnInstructions_Click()
    frmInstructions.Show
End Sub



Private Sub Worksheet_Activate()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayCommentIndicator = xlNoIndicator
    Application.DisplayAlerts = True
    Call MaintenanceAndRepair
End Sub

Sub Worksheet_Change(ByVal Target As Range)
    Dim intTimeInput%
    Dim dblTimeInput#
    Dim strTimeInput$
    Dim tTime As Date
    Dim i%
    Dim dblTimeOutput#
    Dim vNew
    Dim vOld
    Dim intLenStrTimeInput%
    Dim rngWorkingCell As Range

    Application.ScreenUpdating = False
    If Target.Count > 1 Then GoTo EndOfSub
    If Not Intersect(Target, Range("B3:G9")) Is Nothing Then    'checks that the edited cell is in the time punch range
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
        'pull old value to compare later
        Set rngWorkingCell = Range(Target.Address)
        vNew = Range(Target.Address).Value
        Application.EnableEvents = False
        On Error Resume Next
        vOld = rngWorkingCell.Comment.Text
        On Error GoTo 0
        Application.DisplayCommentIndicator = xlNoIndicator
        Application.EnableEvents = True

        On Error Resume Next
        If Target.Value = "" Or Target.Value = " " Or Target.Value = 0 Or Target.Value = Empty Then
            Application.EnableEvents = False
            Target.Value = ""
            Application.EnableEvents = True
            GoTo EndOfSub    'kills if the user just deleted or blanked a cell within the range
        End If
        intTimeInput = Target.Value
        strTimeInput = Target.Value
        dblTimeInput = Target.Value
        intLenStrTimeInput = Len(strTimeInput)
        On Error GoTo 0
        If dblTimeInput < 1 And dblTimeInput > 0 Then dblTimeOutput = dblTimeInput: GoTo AdjustorCode
        If InStr(1, strTimeInput, ".") Then
            strTimeInput = Replace(strTimeInput, ".", ":")
            If (Len(strTimeInput) - InStr(1, strTimeInput, ":") < 2) And (IsNumeric(Right(strTimeInput, 1))) Then strTimeInput = strTimeInput & "0"
            dblTimeInput = CDbl(Replace(strTimeInput, ":", ""))
            intTimeInput = CInt(Replace(strTimeInput, ":", ""))
        End If


        For i = 1 To intLenStrTimeInput
            If Not (i > intLenStrTimeInput) Then
                If IsNumeric(Mid(strTimeInput, i, 1)) = False Then
                    If i = 1 Then
                        strTimeInput = Right(strTimeInput, Len(strTimeInput) - 1)
                        i = 0
                        intLenStrTimeInput = Len(strTimeInput)
                    Else
                        strTimeInput = Trim(Left(strTimeInput, i - 1) & Right(strTimeInput, Len(strTimeInput) - i))
                        i = 0
                        intLenStrTimeInput = Len(strTimeInput)
                    End If
                End If
            Else: i = 100
            End If
        Next i
        If Left(strTimeInput, 1) = 0 Then
            strTimeInput = Trim(Mid(strTimeInput, 2))
            dblTimeInput = CDbl(strTimeInput)
            intTimeInput = CInt(strTimeInput)
        End If

        If strTimeInput = "" Then
            Application.EnableEvents = False
            Target.Value = ""
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
                strM = Mid(strTimeInput, 2, 2)
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



AdjustorCode:
        If Target.Column <> 2 Then
            Dim dblRowMax#
            dblRowMax = Application.WorksheetFunction.Max(Range(Cells(Target.Row, 2), Cells(Target.Row, Target.Column)))
            If dblTimeOutput <= dblRowMax Or (Target.Column = 5 And dblTimeOutput >= dblRowMax And dblTimeOutput < 0.417) Then             ' only auto-PMs OUT times less than 10:00
                dblTimeOutput = dblTimeOutput + 0.5
            End If
        End If
        Application.EnableEvents = False
        On Error Resume Next

        If Round(CDbl(vOld), 4) <> Round(dblTimeOutput, 4) Then
            Range(Target.Address).Interior.ThemeColor = -4142
        Else
            'Range(Target.Address).Interior.ThemeColor = 10
            With Range(Target.Address).Interior    'shade red
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End If
        Range(Target.Address) = dblTimeOutput
    End If














    Dim boolShadedCellsExist As Boolean
    Dim j%
    Dim curValOfCell

EndOfSub:
    Application.EnableEvents = False
    boolShadedCellsExist = False
    For i = 3 To 9
        For j = 2 To 7
            curValOfCell = Cells(i, j).Value
            If Cells(i, j).Value = "" Or Cells(i, j).Value = Empty Then
                With Cells(i, j).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
            If Cells(i, j).Interior.ThemeColor = xlThemeColorAccent6 Then
                boolShadedCellsExist = True
                i = 200: j = 200
            End If
        Next j
    Next i
    '    If boolShadedCellsExist Then
    '        With Cells(17, 2)
    '            .Value = "Note: Red shaded cells denote times imported from PeopleSoft"
    '            .Interior.Pattern = xlSolid
    '            .Interior.PatternColorIndex = xlAutomatic
    '            .Interior.ThemeColor = xlThemeColorAccent6
    '            .Interior.TintAndShade = 0.799981688894314
    '            .Interior.PatternTintAndShade = 0
    '        End With
    '    Else
    '        With Cells(17, 2)
    '            .Value = ""
    '            .Interior.Pattern = xlNone
    '            .Interior.TintAndShade = 0
    '            .Interior.PatternTintAndShade = 0
    '        End With
    '    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True


End Sub


'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'    Application.DisplayCommentIndicator = xlNoIndicator
'    Application.DisplayAlerts = True
'End Sub

