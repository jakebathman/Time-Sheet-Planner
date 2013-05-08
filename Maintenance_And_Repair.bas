Attribute VB_Name = "Maintenance_And_Repair"
Option Explicit

Public Sub MaintenanceAndRepair()

'helps restore the sheets to their original states (including formulas)

'restore application settings (runs on each selection change event also)
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayCommentIndicator = xlNoIndicator
Application.DisplayAlerts = True

End Sub
