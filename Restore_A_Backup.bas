Attribute VB_Name = "Restore_A_Backup"
Option Explicit

Public Sub RestoreABackup()
Call MaintenanceAndRepair

Dim vbRestoreBackup
Dim vbRUSure
Dim i As Integer
Dim strCurSheetName As String
Dim intNumberOfSheets As Integer
Dim intSheetNumForBackup As Integer

intNumberOfSheets = ThisWorkbook.Sheets.Count

For i = 1 To intNumberOfSheets
    strCurSheetName = Sheets(i).Name
    If strCurSheetName = "Backup of Time Sheet Planner" Then
        intSheetNumForBackup = i
        i = 100 'exits the loop
    End If
Next i



vbRestoreBackup = MsgBox("Restore a backup sheet as the main sheet?" _
    & vbCrLf & vbCrLf & "Yes: Delete main sheet and RESTORE BACKUP" _
    & vbCrLf & "No: Keep main sheet and DELETE BACKUP." _
    & vbCrLf & "Cancel: Take no action (keep all sheets)." _
     & vbCrLf & vbCrLf & "(WARNING! Any action taken is not reversible.)", vbYesNoCancel, vbQuestion)

Select Case vbRestoreBackup
    Case vbYes
        'Restore backup sheet (delete main)
        vbRUSure = MsgBox("This will delete the sheet " & Chr(34) & Sheets("Time Sheet Planner").Name & Chr(34) & " and restore the backup sheet in its place." _
            & vbCrLf & vbCrLf & "Continue?", vbYesNo, vbCritical)
        If vbRUSure = vbYes Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Time Sheet Planner").Delete
            Sheets("Backup of Time Sheet Planner").Name = "Time Sheet Planner"
            Application.DisplayAlerts = True
        End If
    Case vbNo
        'Delete backup sheet
        vbRUSure = MsgBox("This will delete the sheet " & Chr(34) & Sheets(intSheetNumForBackup).Name & Chr(34) & " and leave the main sheet alone." _
            & vbCrLf & vbCrLf & "Continue?", vbYesNo, vbCritical)
        If vbRUSure = vbYes Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Backup of Time Sheet Planner").Delete
            Application.DisplayAlerts = True
        End If
    Case vbCancel
        'take no action
End Select

Call MaintenanceAndRepair


End Sub
