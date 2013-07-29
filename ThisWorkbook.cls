VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayCommentIndicator = xlNoIndicator
    Application.DisplayAlerts = True
    Call MaintenanceAndRepair
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    
    If InStr(1, sh.Name, "Time Sheet Planner", vbTextCompare) = 0 Then Exit Sub
    
Dim i%, j%, intTotCol%, intTotRow%
Dim rngCurRange As Range, rngTotCell As Range
Dim i1, o1, i2, o2, i3, o3, timeoff

    '   Formula for total of a row:
        '   2 Punches:  (OUT - IN) - (INLUNCH - OUTLUNCH) + (OUT - IN)
        '   4 Punches:
        '       Lunch:  (OUT - IN) - (INLUNCH - OUTLUNCH)
        '       Night:  (OUT - IN) + (OUT - IN)
        '   6 Punches:  (OUT - IN) - (INLUNCH - OUTLUNCH) + (OUT - IN)

If Intersect(Target, Range("B3:H9")) Is Nothing Then Exit Sub    'checks that the edited cell is in the time punch range

For i = 3 To 9
    i1 = sh.Cells(i, 2).Value
    o1 = sh.Cells(i, 3).Value
    i2 = sh.Cells(i, 4).Value
    o2 = sh.Cells(i, 5).Value
    i3 = sh.Cells(i, 6).Value
    o3 = sh.Cells(i, 7).Value
    timeoff = sh.Cells(i, 8).Value
    Call fCalcPunchTotal(sh, i, i1, o1, i2, o2, i3, o3, timeoff)
Next i

    
End Sub
