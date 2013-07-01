VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInstructions 
   Caption         =   "Instructions"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   OleObjectBlob   =   "frmInstructions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
    Range("B3").Select
End Sub
Private Sub UserForm_Activate()
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub


