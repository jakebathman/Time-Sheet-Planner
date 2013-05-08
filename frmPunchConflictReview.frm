VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPunchConflictReview 
   Caption         =   "Review Existing Punch Conflicts"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   OleObjectBlob   =   "frmPunchConflictReview.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPunchConflictReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer


Private Sub btnUseExisting_Click()
Dim strTempName As String
For i = 1 To 24 Step 2
    strTempName = "OptionButton" & i
    Me.Controls(strTempName).Value = True
Next i
End Sub


Private Sub btnUsePeopleSoft_Click()
Dim strTempName As String
For i = 2 To 24 Step 2
    strTempName = "OptionButton" & i
    Me.Controls(strTempName).Value = True
Next i
End Sub


Private Sub btnUseSelections_Click()

Dim intBtnCounter As Integer
Dim strBtnNameEx As String
Dim strBtnNameSo As String
Dim intNumSelections As Integer
Dim strTempName As String
Dim boolGoodSelections As Boolean
Dim intEscape As Integer

intBtnCounter = 1
intNumSelections = 0
boolGoodSelections = False
intEscape = 1

While boolGoodSelections <> True
intBtnCounter = 1
intNumSelections = 0
intEscape = 1
    If intEscape = 5 Then boolGoodSelections = True
    For i = 1 To intNumTrueConflicts * 2
        strTempName = "OptionButton" & i
        If (Me.Controls(strTempName).Value = True) Then
            intNumSelections = intNumSelections + 1
        End If
    Next i
    Me.Hide
    
    If intNumSelections < intNumTrueConflicts Then
        MsgBox ("You need to make selections for all lines!")
        Me.Show
    Else
        boolGoodSelections = True
    
    
        For i = 1 To intAllArrayLengths
            strBtnNameEx = "OptionButton" & intBtnCounter
            strBtnNameSo = "OptionButton" & (intBtnCounter + 1)
            If arrDoConflictsExist(i) = True Then
                If Me.Controls(strBtnNameEx).Value = True Then
                    arrFinalPunchesToUse(i, 1) = arrExistingForm(i)
                    arrFinalPunchesToUse(i, 2) = "E"
                ElseIf Me.Controls(strBtnNameSo).Value = True Then
                    arrFinalPunchesToUse(i, 1) = arrSortedForm(i)
                ElseIf (Me.Controls(strBtnNameSo).Value = False) And (Me.Controls(strBtnNameEx).Value = False) Then
                    arrFinalPunchesToUse(i, 1) = 0
                End If
                intBtnCounter = intBtnCounter + 2
            Else
                arrFinalPunchesToUse(i, 1) = arrSortedForm(i)
            End If
        Next i
        For i = intAllArrayLengths To UBound(arrFinalPunchesToUse)
            arrFinalPunchesToUse(i, 1) = ""
        Next i
    End If
    intEscape = intEscape + 1
Wend

'Me.Hide

End Sub




'Declared in Module2
    'Public intNumTrueConflicts As Integer
    'Public arrSortedForm() As Variant
    'Public arrExistingForm() As Variant
    'Public arrDoConflictsExist() As Variant
    'Public arrFinalPunchesToUse() As Variant
    'Public intAllArrayLengths As Integer
    'Public arrConflictLocations() As Variant


Public Sub UserForm_Initialize()

Dim i As Integer

Dim strTempName As String
For i = 1 To 24
    strTempName = "OptionButton" & i
    Me.Controls(strTempName).Visible = False
Next i




'make everything visible that needs to be, using if groupname = "groConflict" & chr(65) 'A

Dim strBtnName As String
Dim strBtnNameEx As String
Dim strBtnNameSo As String
Dim intCounterInt As Integer
Dim intBtnCounter As Integer
Dim strSoCap As String
Dim strExCap As String
Dim strPunchDay As String
Dim strPunchStatus As String
Dim intLastBtnTop As Integer

If intNumTrueConflicts <> 0 Then
    For i = 1 To (intNumTrueConflicts * 2)
        strBtnName = "OptionButton" & i
        Me.Controls(strBtnName).Visible = True
    Next i
    intLastBtnTop = Me.Controls(strBtnName).Top
Else
    intLastBtnTop = 50
End If
frmPunchConflictReview.Height = 175 + intLastBtnTop
btnUseSelections.Top = 115 + intLastBtnTop
lblInstructions.Top = 60 + intLastBtnTop
lblInstructionsTitle.Top = 35 + intLastBtnTop


'find the conflicts (TRUE values in array) and make an array with just those locations
intCounterInt = 1
intBtnCounter = 1
If intNumTrueConflicts <> 0 Then
For i = 1 To intAllArrayLengths
    If arrDoConflictsExist(i) = True Then
        ReDim Preserve arrConflictLocations(1 To intCounterInt)
        arrConflictLocations(intCounterInt) = i
        
        strBtnNameEx = "OptionButton" & intBtnCounter
        strBtnNameSo = "OptionButton" & (intBtnCounter + 1)
        strPunchDay = FindDayOfPunch(i)
        strPunchStatus = FindStatusOfPunch(i)
        strExCap = strPunchDay & " " & strPunchStatus & " @ " & arrExistingForm(arrConflictLocations(intCounterInt))
        strSoCap = strPunchDay & " " & strPunchStatus & " @ " & arrSortedForm(arrConflictLocations(intCounterInt))
        Me.Controls(strBtnNameEx).Caption = strExCap
        Me.Controls(strBtnNameSo).Caption = strSoCap

        intBtnCounter = intBtnCounter + 2
        intCounterInt = intCounterInt + 1
    End If
Next i
Else
    With lblNoConflicts
        .Font.Bold = True
        .Font.Size = 10
        .TextAlign = fmTextAlignCenter
        .Visible = True
        .Top = 60
    End With
    btnUseExisting.Visible = False
    btnUsePeopleSoft.Visible = False
    frmPunchConflictReview.Height = 200
    lblInstructions.Visible = False
    lblInstructionsTitle.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    btnUseSelections.Caption = "Continue"
    btnUseSelections.Top = 120
End If

End Sub



Private Sub UserForm_Activate()
Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub






