Attribute VB_Name = "FindSheetInfo"
Option Explicit
Public intSheetNumForBackup As Integer
Public intSheetNumForMain As Integer
Public intNumberOfSheets As Integer
Public strSheetNameForMain As String
Public strSheetNameForBackup As String



Public Sub FindTheSheetInfo()

Dim strCurSheetName As String
Dim i As Integer


intNumberOfSheets = ThisWorkbook.Sheets.Count

intSheetNumForBackup = 42
intSheetNumForMain = 42

'main sheet
For i = 1 To intNumberOfSheets
    strCurSheetName = Sheets(i).Name
    frmPickSheetMain.boxListOfSheetsMain.AddItem strCurSheetName
    If strCurSheetName = "Time Sheet Planner" Then
        intSheetNumForMain = i
    End If
Next i

'backup sheet
For i = 1 To intNumberOfSheets
    strCurSheetName = Sheets(i).Name
    frmPickSheetBackup.boxListOfSheetsBackup.AddItem strCurSheetName
    If strCurSheetName = "Backup of Time Sheet Planner" Then
        intSheetNumForBackup = i
    End If
Next i

If intSheetNumForBackup = 42 Then
    'frmPickSheetBackup.Show
    strSheetNameForBackup = "DOESNOTEXIST"
    On Error Resume Next
    intSheetNumForBackup = Sheets(strSheetNameForBackup).Index
End If
On Error GoTo 0

If intSheetNumForMain = 42 Then
    frmPickSheetMain.Show
    intSheetNumForMain = Sheets(strSheetNameForMain).Index
End If
    


End Sub
