VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNamePicker 
   Caption         =   "What's your name?"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   OleObjectBlob   =   "frmNamePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNamePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    frmNamePicker.cmbEmployeeName.Value = "Choose name . . ."
    Me.Hide
End Sub

Private Sub btnContinue_Click()
    Dim i%, j%
    Dim intCurrentEmployeeColumn%, intRememberCurrentEmployeeColumn%
    Dim boolRememberCurrentEmployee As Boolean
    Dim strCurrentEmployee$

    If frmNamePicker.cmbEmployeeName.ListIndex > 0 Then
        strCurrentEmployee = frmNamePicker.cmbEmployeeName.Value

        With Sheets("References")

            For j = 1 To 10
                If .Cells(1, j).Value = "CurrentEmployee" Then intCurrentEmployeeColumn = j
                If .Cells(1, j).Value = "RememberCurrentEmployee" Then intRememberCurrentEmployeeColumn = j
                If intCurrentEmployeeColumn > 0 And intRememberCurrentEmployeeColumn > 0 Then Exit For
            Next j

            .Cells(2, intCurrentEmployeeColumn).Value = strCurrentEmployee
            .Cells(2, intRememberCurrentEmployeeColumn).Value = frmNamePicker.chkRememberEmployeeName.Value
        End With

        frmNamePicker.Hide

        Sheets("Time Off Form").Activate

    End If

End Sub

Private Sub cmbEmployeeName_Change()
    frmNamePicker.cmbEmployeeName.SelLength = 0
    frmNamePicker.btnContinue.SetFocus
End Sub

Private Sub UserForm_Activate()
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub



Private Sub UserForm_Initialize()
    Dim i%, j%, c%
    Dim intCurrentEmployeeColumn%, intRememberCurrentEmployeeColumn%, intEmployeeNamesColumn%
    Dim boolRememberCurrentEmployee As Boolean
    Dim strCurrentEmployee$
    With Sheets("References")

        For j = 1 To 10
            If .Cells(1, j).Value = "CurrentEmployee" Then intCurrentEmployeeColumn = j
            If .Cells(1, j).Value = "RememberCurrentEmployee" Then intRememberCurrentEmployeeColumn = j
            If .Cells(1, j).Value = "EmployeeNames" Then intEmployeeNamesColumn = j
            If intCurrentEmployeeColumn > 0 And intRememberCurrentEmployeeColumn > 0 And intEmployeeNamesColumn > 0 Then Exit For
        Next j

        strCurrentEmployee = .Cells(2, intCurrentEmployeeColumn).Value
        boolRememberCurrentEmployee = CBool(.Cells(2, intRememberCurrentEmployeeColumn).Value)

        c = 0
        For i = 2 To 25
            If .Cells(i, intEmployeeNamesColumn).Value = "" Then Exit For
            c = c + 1
            Call frmNamePicker.cmbEmployeeName.AddItem(.Cells(i, intEmployeeNamesColumn).Value, i - 2)
            frmNamePicker.cmbEmployeeName.ListRows = c
        Next i

        frmNamePicker.chkRememberEmployeeName.Value = boolRememberCurrentEmployee
        If boolRememberCurrentEmployee And strCurrentEmployee <> "" Then frmNamePicker.cmbEmployeeName.Value = strCurrentEmployee

    End With
    If strCurrentEmployee = "" Or boolRememberCurrentEmployee = False Then frmNamePicker.cmbEmployeeName.ListIndex = 0

    frmNamePicker.cmbEmployeeName.SelLength = 0
   frmNamePicker.cmbEmployeeName.SetFocus
    
End Sub
