VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)

    Dim intTargetRow As Integer
    Dim strTarget As String

    Application.EnableEvents = False

    If Not Intersect(Target, Range("C3:C5")) Is Nothing Then    'checks that the edited cell is in the correct input range
        intTargetRow = Target.Row
        strTarget = Target.Value
        If WorksheetFunction.CountA(Range("C3:C5")) = 0 Then
            Range(Target.Address) = "X"
        ElseIf strTarget <> "" Then
            Select Case intTargetRow
                Case 3
                    Cells(Target.Row, Target.Column) = "X"
                    Cells(Target.Row + 1, Target.Column) = ""
                    Cells(Target.Row + 2, Target.Column) = ""
                Case 4
                    Cells(Target.Row, Target.Column) = "X"
                    Cells(Target.Row + 1, Target.Column) = ""
                    Cells(Target.Row - 1, Target.Column) = ""
                Case 5
                    Cells(Target.Row, Target.Column) = "X"
                    Cells(Target.Row - 1, Target.Column) = ""
                    Cells(Target.Row - 2, Target.Column) = ""
            End Select
        End If
    End If

    Application.EnableEvents = True


End Sub

