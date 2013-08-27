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

Private Sub Worksheet_Activate()
    Dim i%
    Dim v

    With Sheets("References")
        For i = 2 To 50
            If .Cells(i, 1).Value = vbNullString Then
                Sheets("Time Off Form").boxEmployeeName.ListFillRange = "References!$A$2:$A$" & i - 1
                Exit For
            End If
        Next i
    End With
End Sub

