Attribute VB_Name = "m_show_Hide_Time_Off_Buttons"
Option Explicit

Public Sub showHideTimeOffButtons()
    '
    '   BEGIN
    '
    '   Show/hide "Create Time Off Sheet" button
    Dim strTotalTimeOff$, strPTOTime$, strCompTime$, strOtherTimeOff$, strHolidayTime$
    Dim strEmployeeName$
    Dim dblCompAccrued#, dblRateComp#

    dblRateComp = Sheets("User Preferences").Range("B7").Value

    dblCompAccrued = 0
    strTotalTimeOff = "0"
    strPTOTime = "0"
    strCompTime = "0"
    strOtherTimeOff = "0"
    strHolidayTime = "0"

    With Sheets("Time Sheet Planner")
        If .Range("I11").Value <> "" And .Range("I11").Value <> "?" Then strTotalTimeOff = Trim(Mid(.Range("I11").Value, 1, InStr(1, .Range("I11").Value, " ", vbTextCompare)))
        If .Range("I12").Value <> "" And .Range("I12").Value <> "?" Then strPTOTime = .Range("I12").Value
        If .Range("I13").Value <> "" And .Range("I13").Value <> "?" Then strCompTime = .Range("I13").Value
        If .Range("I14").Value <> "" And .Range("I14").Value <> "?" Then strHolidayTime = .Range("I14").Value
        If .Range("I15").Value <> "" And .Range("I15").Value <> "?" Then strOtherTimeOff = .Range("I15").Value

        If .Range("L10").Value <> 0 And .Range("L10").Value > .Range("B1").Value Then dblCompAccrued = (.Range("L10").Value - .Range("B1").Value) * dblRateComp

        If (CDbl(strTotalTimeOff) - CDbl(strHolidayTime) > 0) Then
            .btnCreateTimeOffSheet.Visible = True
            .btnCreateCompForm.Visible = False
        ElseIf (dblCompAccrued > 0) Then
            .btnCreateTimeOffSheet.Visible = False
            .btnCreateCompForm.Visible = True
        Else
            .btnCreateTimeOffSheet.Visible = False
            .btnCreateCompForm.Visible = False
        End If
    End With
End Sub


