Attribute VB_Name = "Module2"
Option Explicit

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
        Operator:=xlBetween, Formula1:="=References!$B$2:$B$5"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = False
        .ShowError = False
    End With
End Sub
