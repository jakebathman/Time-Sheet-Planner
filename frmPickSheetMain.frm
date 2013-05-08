VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickSheetMain 
   Caption         =   "Pick the correct main sheet"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
   OleObjectBlob   =   "frmPickSheetMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPickSheetMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    End
End Sub

Public Sub btnUseThisSheet_Click()
    strSheetNameForMain = frmPickSheetMain.boxListOfSheetsMain.Value
    Me.Hide
End Sub

Private Sub UserForm_Activate()
Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub
