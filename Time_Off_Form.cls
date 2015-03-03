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
Public boolChangingStuff As Boolean

Private Sub boxEmployeeName_Change()
    If boxEmployeeName.Value <> "Choose name . . ." Then
        boxEmployeeName.BackColor = &HFFFFFF
    Else
        boxEmployeeName.BackColor = &HC0E0FF
    End If
    boxEmployeeName.SelLength = 0
    Range("C2:D2").Select
End Sub

Private Sub btnPrintTimeOffForm_Click()
    ' Just used for debugging:
    'ActiveSheet.PrintPreview
    
    Sheets("Time Off Form").PrintOut From:=1, To:=1
End Sub

Public Sub btnResetTimeOffForm_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    boxEmployeeName.Value = "Choose name . . ."
    Range("D4").Value = ""  ' from time
    Range("D6").Value = ""  ' to time
    Range("C4").Value = ""  ' from date
    Range("C6").Value = ""  ' to date
    Range("H6").Value = ""  ' total hrs

    Range("H2").Value = Date   ' sets Date Submitted to today's date


    boolChangingStuff = True
    Call ChkAllFalse
    boolChangingStuff = False
    Call ColorNeededChks

    Range("A24:J30").ClearContents    ' Reason text
    Range("D32:I32").ClearContents    ' Reason text comp earned
    Range("D35:I35").ClearContents    ' Reason text overtime earned
    Call ColorNeededChks
    Range("C2:D2").Select    ' total hrs
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Sub chkCat_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkCat.Value

    Call ChkAllFalse

    boolChangingStuff = True
    chkCat = boolThisCheck
    boolChangingStuff = False
    Call ColorNeededChks
End Sub

Private Sub chkComp_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkComp.Value

    Call ChkAllFalse

    boolChangingStuff = True
    chkComp.Value = boolThisCheck
    boolChangingStuff = False
    Call ColorNeededChks
End Sub

Private Sub chkCompEarned_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkCompEarned.Value

    Call ChkAllFalse

    boolChangingStuff = True
    chkCompEarned.Value = boolThisCheck
    boolChangingStuff = False
    Call ColorNeededChks
End Sub

Private Sub chkOther_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkOther.Value

    Call ChkAllFalse

    boolChangingStuff = True
    chkOther.Value = boolThisCheck
    boolChangingStuff = False
    Call ColorNeededChks
End Sub

Private Sub chkOvertimeEarned_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkOvertimeEarned.Value

    Call ChkAllFalse

    boolChangingStuff = True
    chkOvertimeEarned.Value = boolThisCheck
    boolChangingStuff = False
    Call ColorNeededChks
End Sub



Private Sub chkPTO_Change()
    Dim boolThisCheck As Boolean
    If boolChangingStuff Then Exit Sub

    boolThisCheck = Me.chkPTO.Value

    Call ChkAllFalse
    boolChangingStuff = True
    chkPTO.Value = boolThisCheck
    boolChangingStuff = False

    Call ColorNeededChks
End Sub

Private Sub Worksheet_Activate()
    Dim i%
    Dim v

    ' My code
    With Sheets("References")
        For i = 2 To 50
            If .Cells(i, 1).Value = vbNullString Then
                Sheets("Time Off Form").boxEmployeeName.ListFillRange = "References!$A$2:$A$" & i - 1
                Sheets("Time Off Form").boxEmployeeName.ListRows = i
                Exit For
            End If
        Next i
    End With

End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If ActiveWindow.View <> xlNormalView Then ActiveWindow.View = xlNormalView
    If Target.Address = "$C$4" Or Target.Address = "$C$6" Or Target.Address = "$H$2" Then
        CalendarFrm.Show
        Unload CalendarFrm
    End If

End Sub




Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False

    If Range("$C$6").Value < Range("$C$4").Value Then Range("$C$6").Value = Range("$C$4").Value

    ' Color the reason for time off box if it's blank
    If chkPTO.Value = True Or chkCat.Value = True Or chkComp.Value = True Or chkOther.Value = True Then
        If Range("A24").Value = vbNullString Then
            Call MainReasonOrange
        Else
            Call MainReasonWhite
        End If
    End If


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

    Call ColorNeededChks

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub


End Sub

Public Sub ColorNeededChks()
    Application.EnableEvents = False

    With Sheets("Time Off Form")
        If chkPTO.Value = False And chkComp.Value = False And chkCat.Value = False And chkOther.Value = False And chkOvertimeEarned.Value = False And chkCompEarned.Value = False Then
            ' orange
            With .Range("B10").Interior
                .Pattern = xlUp
                .PatternThemeColor = xlThemeColorDark1
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.399945066682943
                .PatternTintAndShade = 0
            End With
            With .Range("B12").Interior
                .Pattern = xlUp
                .PatternThemeColor = xlThemeColorDark1
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.399945066682943
                .PatternTintAndShade = 0
            End With
            With .Range("B19").Interior
                .Pattern = xlUp
                .PatternThemeColor = xlThemeColorDark1
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.399945066682943
                .PatternTintAndShade = 0
            End With
            With .Range("B21").Interior
                .Pattern = xlUp
                .PatternThemeColor = xlThemeColorDark1
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.399945066682943
                .PatternTintAndShade = 0
            End With
        Else
            .Range("B10").ClearFormats
            .Range("B12").ClearFormats
            .Range("B19").ClearFormats
            .Range("B21").ClearFormats
        End If

        If chkCompEarned.Value = True Then
            Call MainReasonWhite
            If .Range("D32").Value = vbNullString Then
                'orange
                With .Range("D32:I32").Interior
                    .Pattern = xlUp
                    .PatternThemeColor = xlThemeColorDark1
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.399945066682943
                    .PatternTintAndShade = 0
                End With
            Else
                'white
                With .Range("D32:I32").Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Else
            'white
            With .Range("D32:I32").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

        End If

        If chkOvertimeEarned.Value = True Then
            Call MainReasonWhite
            If .Range("D35").Value = vbNullString Then
                'orange
                With .Range("D35:I35").Interior
                    .Pattern = xlUp
                    .PatternThemeColor = xlThemeColorDark1
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.399945066682943
                    .PatternTintAndShade = 0
                End With
            Else
                'white
                With .Range("D35:I35").Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Else
            'white
            With .Range("D35:I35").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

        End If




        ' Color the reason for time off box if it's blank
        If chkPTO.Value = True Or chkCat.Value = True Or chkComp.Value = True Or chkOther.Value = True Then
            If Range("A24").Value = vbNullString Then
                Call MainReasonOrange
            End If
        Else
            Call MainReasonWhite
        End If



    End With
    Application.EnableEvents = True

End Sub


Public Sub ChkAllFalse()
    boolChangingStuff = True
    chkPTO.Value = False
    chkComp.Value = False
    chkCat.Value = False
    chkOther.Value = False
    chkCompEarned.Value = False
    chkOvertimeEarned.Value = False
    boolChangingStuff = False
End Sub

Public Sub MainReasonWhite()
    'white
    With Range("A24:I24").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Public Sub MainReasonOrange()
    Range("A24:I24").ClearContents
    With Range("A24:I24").Interior
        .Pattern = xlUp
        .PatternThemeColor = xlThemeColorDark1
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
        .PatternTintAndShade = 0
    End With
End Sub
