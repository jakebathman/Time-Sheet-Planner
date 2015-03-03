VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickTimeOffCode 
   Caption         =   "Which time off code?"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmPickTimeOffCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPickTimeOffCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCancel_Click()
    frmPickTimeOffCode.cmbPickTimeOffCode.Value = "Pick one..."
    frmPickTimeOffCode.Hide
End Sub

Private Sub btnContinue_Click()
    If frmPickTimeOffCode.cmbPickTimeOffCode.ListIndex > 0 Then
        frmPickTimeOffCode.Hide
    End If
End Sub

Private Sub cmbPickTimeOffCode_Change()

    frmPickTimeOffCode.cmbPickTimeOffCode.SelLength = 0
    frmPickTimeOffCode.btnContinue.SetFocus
End Sub

Private Sub UserForm_Activate()
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub



Private Sub UserForm_Initialize()
    Dim strTotalTimeOff$, strPTOTime$, strCompTime$, strOtherTimeOff$, strHolidayTime$, strClosureTime$
    Dim strEmployeeName$
    Dim boolMultipleTimeOffCodes As Boolean
    Dim intCountTimeOffCodes%

    strTotalTimeOff = "0"
    strPTOTime = "0"
    strCompTime = "0"
    strOtherTimeOff = "0"
    strClosureTime = "0"
    strHolidayTime = "0"


    With Sheets("Time Sheet Planner")
        If .Range("I11").Value <> "" And .Range("I11").Value <> "?" Then strTotalTimeOff = Trim(Mid(.Range("I11").Value, 1, InStr(1, .Range("I11").Value, " ", vbTextCompare)))
        If .Range("I12").Value <> "" And .Range("I12").Value <> "?" Then strPTOTime = .Range("I12").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
        If .Range("I13").Value <> "" And .Range("I13").Value <> "?" Then strCompTime = .Range("I13").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
        If .Range("I14").Value <> "" And .Range("I14").Value <> "?" Then strHolidayTime = .Range("I14").Value
        If .Range("I15").Value <> "" And .Range("I15").Value <> "?" Then strClosureTime = .Range("I15").Value
        If .Range("I16").Value <> "" And .Range("I16").Value <> "?" Then strOtherTimeOff = .Range("I16").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
    End With

    frmPickTimeOffCode.cmbPickTimeOffCode.AddItem "Pick one..."
    If CDbl(strPTOTime) > 0 Then frmPickTimeOffCode.cmbPickTimeOffCode.AddItem ("PTO - " & CDbl(strPTOTime) & " hrs")
    If CDbl(strCompTime) > 0 Then frmPickTimeOffCode.cmbPickTimeOffCode.AddItem ("Comp - " & CDbl(strCompTime) & " hrs")
    If CDbl(strOtherTimeOff) > 0 Then frmPickTimeOffCode.cmbPickTimeOffCode.AddItem ("Other - " & CDbl(strOtherTimeOff) & " hrs")

    frmPickTimeOffCode.cmbPickTimeOffCode.ListIndex = 0
End Sub
