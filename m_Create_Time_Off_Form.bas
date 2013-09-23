Attribute VB_Name = "m_Create_Time_Off_Form"
Option Explicit

Public Sub CreateTimeOffForm()
    Dim strTotalTimeOff$, strPTOTime$, strCompTime$, strOtherTimeOff$, strHolidayTime$
    Dim strEmployeeName$, strTimeOffCodeToUse$
    Dim boolMultipleTimeOffCodes As Boolean
    Dim intCountTimeOffCodes%

    strTotalTimeOff = "0"
    strPTOTime = "0"
    strCompTime = "0"
    strOtherTimeOff = "0"
    strHolidayTime = "0"


    With Sheets("Time Sheet Planner")
        If .Range("I11").Value <> "" And .Range("I11").Value <> "?" Then strTotalTimeOff = Trim(Mid(.Range("I11").Value, 1, InStr(1, .Range("I11").Value, " ", vbTextCompare)))
        If .Range("I12").Value <> "" And .Range("I12").Value <> "?" Then strPTOTime = .Range("I12").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
        If .Range("I13").Value <> "" And .Range("I13").Value <> "?" Then strCompTime = .Range("I13").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
        If .Range("I14").Value <> "" And .Range("I14").Value <> "?" Then strHolidayTime = .Range("I14").Value
        If .Range("I15").Value <> "" And .Range("I15").Value <> "?" Then strOtherTimeOff = .Range("I15").Value: intCountTimeOffCodes = intCountTimeOffCodes + 1
    End With

    If CDbl(strTotalTimeOff) > 0 Then
        If intCountTimeOffCodes > 1 Then
            frmPickTimeOffCode.Show
            If frmPickTimeOffCode.cmbPickTimeOffCode.Value = "" Or frmPickTimeOffCode.cmbPickTimeOffCode.Value = "Pick one..." Then Exit Sub
            strTimeOffCodeToUse = Mid(frmPickTimeOffCode.cmbPickTimeOffCode.Value, 1, InStr(1, frmPickTimeOffCode.cmbPickTimeOffCode.Value, " ", vbTextCompare) - 1)
            Unload frmPickTimeOffCode
        Else
            If strPTOTime > 0 Then strTimeOffCodeToUse = "PTO"
            If strCompTime > 0 Then strTimeOffCodeToUse = "Comp"
            If strOtherTimeOff > 0 Then strTimeOffCodeToUse = "Other"
            If strTimeOffCodeToUse = "" Then strTimeOffCodeToUse = "Total"
        End If

        frmNamePicker.Show
        If frmNamePicker.cmbEmployeeName.ListIndex = 0 Or frmNamePicker.cmbEmployeeName.Value = "Choose name . . ." Then Unload frmNamePicker: Exit Sub
        strEmployeeName = frmNamePicker.cmbEmployeeName.Value
        Unload frmNamePicker
        
        Sheets("Time Off Form").Activate

        With Sheets("Time Off Form")
            Application.EnableEvents = False

            Call Sheets("Time Off Form").btnResetTimeOffForm_Click

            .boxEmployeeName.Value = strEmployeeName

            Select Case strTimeOffCodeToUse
                Case Is = "PTO"
                    .Range("H6").Value = strPTOTime  ' total hrs
                    .chkPTO.Value = True
                Case Is = "Comp"
                    .Range("H6").Value = strCompTime  ' total hrs
                    .chkComp.Value = True
                Case Is = "Other"
                    .Range("H6").Value = strOtherTimeOff  ' total hrs
                    .chkOther.Value = True
                Case Else
                    .Range("H6").Value = strTotalTimeOff  ' total hrs
                    .chkOther.Value = True
            End Select

            .Range("H2").Value = Date   ' sets Date Submitted to today's date


            .Range("C4").Select

        End With


        Application.EnableEvents = True
    End If
End Sub







