Attribute VB_Name = "f_Calc_Punch_Total"
Option Explicit

Public Function fCalcPunchTotal(ByRef sh As Worksheet, curRow As Integer, in1, out1, in2, out2, in3, out3, Optional timeoff)

    Dim i%, j%, intTotCol%, intTotRow%, intPunchesInRow%
    Dim rngCurRange As Range, rngTotCell As Range
    Dim t1, t2, t3

    '   Formula for total of a row:
    '   2 Punches:  (OUT - IN)
    '   4 Punches:
    '       Lunch:  (OUT - IN) - (INLUNCH - OUTLUNCH)
    '       Night:  (OUT - IN) + (OUT - IN)
    '   6 Punches:  (OUT - IN) - (INLUNCH - OUTLUNCH) + (OUT - IN)

    Set rngCurRange = sh.Range(Cells(curRow, 2), Cells(curRow, 7))
    intPunchesInRow = Application.CountA(rngCurRange)
    If IsError(intPunchesInRow) Then intPunchesInRow = 0
    Set rngTotCell = Range("L" & curRow)

    'rngCurRange.Select

    Application.EnableEvents = False

    ' if only the first two columns have times, change computation method
    If out1 > 0 And in1 > 0 And (out2 = vbNullString And out3 = vbNullString And in2 = vbNullString And in3 = vbNullString) Then
        t1 = (fRoundTime(out1) - fRoundTime(in1)) * 24    ' early day two punches
        If t1 >= 0 Then
            rngTotCell.Value = Format(t1 + timeoff, "#.00")
        Else
            rngTotCell.Value = ""
        End If
    Else
        t1 = (fRoundTime(in2) - fRoundTime(out1)) * 24    ' lunch
        t2 = (fRoundTime(out2) - fRoundTime(in1)) * 24    ' full day (including lunch)
        t3 = (fRoundTime(out3) - fRoundTime(in3)) * 24    ' final in/out set
        If out1 = vbNullString Or in2 = vbNullString Then t1 = 0
        If out2 = vbNullString Or in1 = vbNullString Then t2 = 0
        If out3 = vbNullString Or in3 = vbNullString Then t3 = 0
        Select Case intPunchesInRow
            Case 0, 2, 4, 6
                If ((t2 - t1) + t3 + timeoff) = 0 Then
                    rngTotCell.Value = ""
                Else
                    rngTotCell.Value = Format((t2 - t1) + t3 + timeoff, "#.00")
                End If
            Case Else
                rngTotCell.Value = ""
        End Select
    End If
    Application.EnableEvents = True


End Function

Public Function fRoundTime(ByVal t As Double) As Double
    Dim intHr%, intMin%, intSec%
    Dim h#, m#, s#
    Dim fh, fm, Fs

    If t = 0 Or t >= 1 Then fRoundTime = 0
    h = t * 24
    m = t * 24 * 60
    s = t * 24 * 60 * 60
    fh = Floor(h)
    fm = Floor(m)
    Fs = Floor(s)
    intHr = Floor(h + 0.00001)
    intMin = Floor(((h - fh) * 60) + 0.00001)
    intSec = Floor((m - fm) * 60 + 0.00001)
    'Debug.Print "Hr:  " & intHr
    'Debug.Print "Min: " & intMin
    'Debug.Print "Sec: " & intSec
    'select case for rounding rules
    Select Case intMin
        Case 7, 22, 37, 52
            If intSec <= 29 Then
                intMin = intMin - 7
            Else
                If intMin = 52 Then
                    intHr = intHr + 1
                    intMin = 0
                Else
                    intMin = intMin + 8
                End If
            End If
        Case 0 To 6
            intMin = 0
        Case 8 To 21
            intMin = 15
        Case 23 To 36
            intMin = 30
        Case 38 To 51
            intMin = 45
        Case 53 To 59
            intMin = 0
            intHr = intHr + 1
    End Select
    intSec = 0



    'Debug.Print "Hr:  " & intHr
    'Debug.Print "Min: " & intMin
    'Debug.Print "Sec: " & intSec

    fRoundTime = (intHr / 24) + (intMin / 24 / 60) + (intSec / 24 / 60 / 60)




End Function



Public Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

Public Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Floor = Int(X / Factor) * Factor
End Function
