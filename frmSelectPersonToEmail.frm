VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectPersonToEmail 
   Caption         =   "Select a person to email your time"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   OleObjectBlob   =   "frmSelectPersonToEmail.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSelectPersonToEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolDropFocus As Boolean
Dim strOldEmail$, strNewEmail$, strOldEmailPrefix$
Dim strClearOldEmail$, strClearOldName$

Private Sub boxOtherEmail_AfterUpdate()
    boolDropFocus = False
    Me.boxOtherEmail.SelStart = 0
End Sub

Private Sub boxOtherEmail_Change()
        'Me.chkClear.Enabled = True
        'Me.chkClear.Value = False
    strNewEmail = Me.boxOtherEmail.Value
    boolDropFocus = False
    If strNewEmail = "" Then
        strOldEmail = "": strOldEmailPrefix = "": Me.boxOtherEmail.Value = ""
        'Me.chkClear.Enabled = False
        'Me.chkClear.Value = False
    Else
        If InStr(1, strNewEmail, "@co.collin.tx.us") = 0 And Len(strOldEmail) > Len("@co.collin.tx.us") Then 'fix string missing county suffix
            Me.boxOtherEmail.Value = strOldEmailPrefix & "@co.collin.tx.us"
        ElseIf InStr(1, strNewEmail, "co.collin.tx.us") Then
            Me.boxOtherEmail.Value = Left(Me.boxOtherEmail.Value, Len(Me.boxOtherEmail.Value) - Len("@co.collin.tx.us")) & "@co.collin.tx.us"
        Else
            Me.boxOtherEmail.Value = Me.boxOtherEmail.Value & "@co.collin.tx.us"
        End If
        Me.boxOtherEmail.SelStart = Len(Me.boxOtherEmail.Value) - Len("@co.collin.tx.us")
        strOldEmail = Me.boxOtherEmail.Value
        strOldEmailPrefix = Left(strOldEmail, InStr(1, strOldEmail, "@") - 1)
    End If
End Sub

Private Sub boxOtherEmail_Enter()
    boolDropFocus = False
End Sub

Private Sub boxOtherName_Change()
        'Me.chkClear.Enabled = True
        'Me.chkClear.Value = False
    boolDropFocus = False
    Me.boxOtherName.Value = StrConv(Me.boxOtherName.Value, vbProperCase)
    Dim intSp%, intSpTwo%
    intSp = InStr(1, Me.boxOtherName.Value, " ")
    intSpTwo = InStr(intSp + 1, Me.boxOtherName.Value, " ")
    If intSp = 0 Then
        Me.boxOtherEmail.Value = Left(Me.boxOtherName.Value, 1)
    Else
        If intSpTwo <> 0 Then
            Me.boxOtherEmail.Value = Left(Me.boxOtherName.Value, 1) & Trim$(Mid(Me.boxOtherName.Value, intSp + 1, intSpTwo - intSp)) & "@co.collin.tx.us"
        Else
            Me.boxOtherEmail.Value = Left(Me.boxOtherName.Value, 1) & Trim$(Mid(Me.boxOtherName.Value, intSp + 1)) & "@co.collin.tx.us"
        End If
    End If
    If Me.boxOtherName.Value = "" Then
        Me.boxOtherEmail.Value = ""
        'Me.chkClear.Enabled = False
        'Me.chkClear.Value = False
    End If
End Sub

Private Sub boxOtherName_Enter()
    boolDropFocus = False
End Sub

Private Sub boxPeople_Change()
    boolDropFocus = False
    If Left(Me.boxPeople.Value, 5) = "Other" Or Me.boxPeople.Value = arrPeopleAndEmails(5, 1) Or Me.boxPeople.Value = arrPeopleAndEmails(6, 1) Then
        Me.Height = 170
        Me.btnCancel.Top = 115
        Me.btnSubmit.Top = 115
        Me.boxOtherName.Visible = True
        Me.boxOtherName.Enabled = True
        Me.boxOtherEmail.Visible = True
        Me.boxOtherEmail.Enabled = True
        Me.lblOtherEmail.Visible = True
        Me.lblOtherName.Visible = True
        Me.lblOtherNote.Visible = True
        Me.lblOtherNote.Top = 95
        Me.chkClear.Visible = True
        Me.chkClear.Enabled = True
        Me.chkClear.Value = False
        Me.Label2.Visible = True
        Me.Label3.Visible = True
            If Left(Me.boxPeople.Value, 5) <> "Other" Then
                Select Case Me.boxPeople.Value
                    Case arrPeopleAndEmails(5, 1)
                        Me.boxOtherName.Value = arrPeopleAndEmails(5, 1)
                        Me.boxOtherEmail.Value = arrPeopleAndEmails(5, 2)
                    Case arrPeopleAndEmails(6, 1)
                        Me.boxOtherName.Value = arrPeopleAndEmails(6, 1)
                        Me.boxOtherEmail.Value = arrPeopleAndEmails(6, 2)
                End Select
            End If
        Me.boxOtherName.SetFocus
        boolDropFocus = False
    Else
        Me.Height = 115
        Me.btnCancel.Top = 60
        Me.btnSubmit.Top = 60
        Me.boxOtherName.Visible = False
        Me.boxOtherName.Enabled = False
        Me.boxOtherEmail.Visible = False
        Me.boxOtherEmail.Enabled = False
        Me.lblOtherEmail.Visible = False
        Me.lblOtherName.Visible = False
        Me.lblOtherNote.Visible = False
        'Me.chkClear.Value = False
        Me.chkClear.Visible = False
        Me.Label2.Visible = False
        Me.Label3.Visible = False
        Me.btnSubmit.SetFocus
        boolDropFocus = False
    End If
End Sub

Private Sub boxPeople_Click()
    If Left(Me.boxPeople.Value, 5) <> "Other" Then Me.btnSubmit.SetFocus
    'boolDropFocus = True
End Sub

Private Sub boxPeople_Enter()
    Me.boxPeople.DropDown
    boolDropFocus = True
End Sub


Private Sub btnSubmit_Click()
boolDropFocus = False
boolDone = False
Application.EnableEvents = False
Me.Hide
On Error Resume Next
    If Me.chkClear.Value = True Then
        If strClearOldName = arrPeopleAndEmails(5, 1) Then
            Sheets("User Preferences").Cells(intOtherEmailsFirstRow, 2) = ""
            Sheets("User Preferences").Cells(intOtherEmailsFirstRow, 3) = ""
        End If
        If strClearOldName = arrPeopleAndEmails(6, 1) Then
            Sheets("User Preferences").Cells(intOtherEmailsFirstRow + 1, 2) = ""
            Sheets("User Preferences").Cells(intOtherEmailsFirstRow + 1, 3) = ""
        End If
    End If

While Not boolDone
    Select Case Me.boxPeople
        Case "Eileen P."
            strEmail = "eprentice@co.collin.tx.us"
            strName = Me.boxPeople.Value
            strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
            boolDone = True
        Case "Lawana D."
            strEmail = "ldowns@co.collin.tx.us"
            strName = Me.boxPeople.Value
            strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
            boolDone = True
        Case "Carol S."
            boolDone = True
            strName = Me.boxPeople.Value
            strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
            strEmail = "cstrickland@co.collin.tx.us"
        Case "Jake B."
            boolDone = True
            strName = "Jake"
            strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
            strEmail = "jbathman@co.collin.tx.us"
        Case arrPeopleAndEmails(5, 1)
            If Me.boxOtherEmail.Value <> "" And Me.boxOtherName.Value <> "" Then
                boolDone = True
                strName = Me.boxOtherName.Value
                strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
                strEmail = Trim(Me.boxOtherEmail.Value)
                Sheets("User Preferences").Cells(intOtherEmailsFirstRow, 2) = strName
                Sheets("User Preferences").Cells(intOtherEmailsFirstRow, 3) = strEmail
            Else
                boolDone = False
                If MsgBox("Sorry, that wasn't a valid selection. Please select from the list.", vbOKCancel) = vbCancel Then End
                Me.Show
            End If
        Case arrPeopleAndEmails(6, 1)
            If Me.boxOtherEmail.Value <> "" And Me.boxOtherName.Value <> "" Then
                boolDone = True
                strName = Me.boxOtherName.Value
                strName = Left(strName, InStr(1, strName, " ", vbTextCompare) - 1)
                strEmail = Trim(Me.boxOtherEmail.Value)
                Sheets("User Preferences").Cells(intOtherEmailsFirstRow + 1, 2) = strName
                Sheets("User Preferences").Cells(intOtherEmailsFirstRow + 1, 3) = strEmail
            Else
                boolDone = False
                If MsgBox("Sorry, that wasn't a valid selection. Please select from the list.", vbOKCancel) = vbCancel Then End
                Me.Show
            End If
        Case Else
            If MsgBox("Sorry, that wasn't a valid selection. Please select from the list.", vbOKCancel) = vbCancel Then End
            Me.Show
    End Select
Wend
Application.EnableEvents = True
On Error GoTo 0
End Sub

Private Sub btnCancel_Click()
    boolDropFocus = False
    End
End Sub


Private Sub chkClear_Click()
    If Me.chkClear.Value = True Then
        strClearOldEmail = Me.boxOtherEmail.Value
        strClearOldName = Me.boxOtherName.Value
        Me.boxOtherEmail.Value = ""
        Me.boxOtherName.Value = ""
    Else
        Me.boxOtherEmail.Value = strClearOldEmail
        Me.boxOtherName.Value = strClearOldName
    End If
End Sub

Private Sub Label3_Click()
    MsgBox ("Check this box to quickly clear inputted values, and to overwrite the names/emails for future use. Leaving the boxes blank, even if selecting another name, will clear the record.")
End Sub

Private Sub UserForm_Click()
    If boolDropFocus Then Me.btnSubmit.SetFocus
    boolDropFocus = False
End Sub

Private Sub UserForm_Initialize()
    boolDropFocus = False
    Call UserForm_Activate
End Sub

Private Sub UserForm_Activate()
    boolDropFocus = False
    Me.boxPeople.List = arrPeople
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub
