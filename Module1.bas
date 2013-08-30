Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWindow.View = xlPageLayoutView
    ActiveWindow.View = xlNormalView
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("D6").Select
    Selection.NumberFormat = "[$-409]h:mm AM/PM;@"
    ActiveWorkbook.Save
End Sub
