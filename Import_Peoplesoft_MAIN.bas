Attribute VB_Name = "Import_Peoplesoft_MAIN"
Option Explicit

Public intNumTrueConflicts As Integer
Public arrSortedForm() As Variant
Public arrExistingForm() As Variant
Public arrDoConflictsExist() As Variant
Public arrFinalPunchesToUse() As Variant
Public intAllArrayLengths As Integer
Public arrConflictLocations() As Variant
Public dblTimeNowRnd#
Public strWindowSearchTitle As String
Public boolFoundWindow As Boolean

Public Sub PeopeSoftImport()

    Call MaintenanceAndRepair

    If MsgBox("WARNING!" & vbNewLine & vbNewLine & "Recent updates to PeopleSoft have broken this importer!" & vbNewLine & _
              "I'd advise you enter your time manually, until Jake updates the code." & vbNewLine & vbNewLine & "Continue anyway??", vbCritical + vbYesNo, "Hold on, there!") = vbNo Then Exit Sub

    dblTimeNowRnd = Now()
    Application.EnableEvents = False


    Dim i As Integer
    Dim j As Integer
    Dim intYesNo
    Dim vbRUSure
    Dim boolBackupOperationComplete As Boolean


    'Begin and prompt to continue.
    strWindowSearchTitle = "Timesheet"
    boolFoundWindow = False
    Call SearchForWindowByTitle

    Select Case boolFoundWindow
        Case False    'Looks like Timesheet is already open
            intYesNo = MsgBox("Automatically open the PeopleSoft website in your browser?" & vbNewLine & vbNewLine _
                            & "Selecting ""No"" will require you to navigate to your timesheet manually.", vbYesNoCancel + vbQuestion)
            Select Case intYesNo
                Case vbYes
                    ThisWorkbook.FollowHyperlink "https://employees.co.collin.tx.us/psp/EMPSS/EMPLOYEE/HRMS/c/ROLE_EMPLOYEE.TL_MSS_EE_SRCH_PRD.GBL?PORTALPARAM_PTCNAV=HC_TL_SS_JOB_SRCH_EE_GBL&EOPP.SCNode=HRMS&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=CO_EMPLOYEE_SELF_SERVICE&EOPP.SCLabel=Report Time&EOPP.SCFName=HC_RECORD_TIME&EOPP.SCSecondary=true&EOPP.SCPTfname=HC_RECORD_TIME&FolderPath=PORTAL_ROOT_OBJECT.CO_EMPLOYEE_SELF_SERVICE.HC_TIME_REPORTING.HC_RECORD_TIME.HC_TL_SS_JOB_SRCH_EE_GBL&IsFolder=false"
                Case vbNo
                    'foo
                Case Else
                    End
            End Select
        Case True
            intYesNo = MsgBox("Looks like you've already got PeopleSoft open." & vbNewLine & vbNewLine _
                            & "Continue import?", vbYesNoCancel + vbQuestion)
            Select Case intYesNo
                Case vbYes
                    'foo
                Case vbNo
                    End
                Case Else
                    End
            End Select
    End Select




    Dim boolOverwriteBackup
    Dim intCurSheetNum As Integer


    'create backups
    boolBackupOperationComplete = False
    Application.DisplayAlerts = False

    On Error Resume Next
    IsError (Sheets("Backup of Time Sheet Planner").Index)

    If Err.Number <> 9 Then
        While boolBackupOperationComplete = False
            boolOverwriteBackup = MsgBox("Backup of main sheet already exists. Overwrite?", vbYesNoCancel)
            Select Case boolOverwriteBackup
                Case vbYes
                    'creates hidden backup sheet of backup, just in case. Not accessed anywhere else, must be manually reinstated
                    On Error Resume Next
                    Sheets("Hidden Backup of Old Backup").Visible = True
                    Sheets("Hidden Backup of Old Backup").Delete
                    On Error GoTo ErrHandlerCode
                    Sheets("Backup of Time Sheet Planner").Activate
                    Sheets("Backup of Time Sheet Planner").Copy after:=Sheets("Backup of Time Sheet Planner")
                    ActiveSheet.Name = "TmpSheet" & dblTimeNowRnd + Rnd
                    intCurSheetNum = ActiveSheet.Index
                    ActiveSheet.Name = "Hidden Backup of Old Backup"
                    ActiveSheet.Visible = xlSheetVeryHidden
                    Sheets("Backup of Time Sheet Planner").Delete
                    Sheets(1).Activate
                    Sheets(1).Copy after:=Sheets(1)
                    ActiveSheet.Name = "Backup of Time Sheet Planner"
                    boolBackupOperationComplete = True
                Case vbNo
                    vbRUSure = MsgBox("This won't back up your current sheet, but keep a (possibly) old backup. Really continue without backing up main sheet?", vbYesNo + vbQuestion)
                    If vbRUSure = vbYes Then boolBackupOperationComplete = True
                Case vbCancel
                    End
            End Select
        Wend
    Else
        Sheets("Time Sheet Planner").Activate
        ActiveSheet.Copy after:=Sheets("Time Sheet Planner")
        ActiveSheet.Name = "Backup of Time Sheet Planner"
    End If

    'creates hidden backup of main sheet, just in case. Not accessed anywhere else, must be manually reinstated
    On Error Resume Next
    Sheets("Hidden Backup of Main").Visible = True
    Sheets("Hidden Backup of Main").Delete
    On Error GoTo ErrHandlerCode
    Sheets("Time Sheet Planner").Activate
    Sheets("Time Sheet Planner").Copy after:=Sheets("Backup of Time Sheet Planner")
    ActiveSheet.Name = "TmpSheet" & dblTimeNowRnd + Rnd
    intCurSheetNum = ActiveSheet.Index
    ActiveSheet.Name = "Hidden Backup of Main"
    ActiveSheet.Visible = xlSheetVeryHidden

    Application.DisplayAlerts = True






    'Make sure sheet called "PeopleSoft" exists, delete it if so, and clear it completely
    Dim intNumWorksheets As Integer
    intNumWorksheets = ThisWorkbook.Sheets.Count

    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("PeopleSoft").Delete
    On Error GoTo ErrHandlerCode
    intNumWorksheets = ThisWorkbook.Sheets.Count
    Sheets.Add(after:=Sheets(intNumWorksheets)).Name = "PeopleSoft"
    Application.DisplayAlerts = True


    'Give user instructions on copying text
    Dim intYesNoPreCopy As Integer
    Dim boolContinueWithPaste As Boolean
    Dim intRUSure As Integer
    Dim intDoneCopying As Integer


    boolContinueWithPaste = False

    While boolContinueWithPaste <> True
        intYesNoPreCopy = MsgBox("Time for your job. Here are the steps to copy content from PeopleSoft (leave this box up and come back when you're done):" & vbNewLine & vbNewLine _
                               & "1. Log into PeopleSoft in your browser" & vbNewLine _
                               & "2. Navigate to Self Service... Time Reporting... Report Time... Timesheet" & vbNewLine _
                               & "3. Click anywhere in the white space below the timesheet boxes, then press Control + a to select the whole timesheet" & vbNewLine _
                               & "4. With everything selected, copy it to the clipboard" & vbNewLine _
                               & "5. After copying to the clipboard, head back to this box and click OK below to continue" & vbNewLine & vbNewLine _
                               & "In the future, you can do this before running the macro.", vbOKCancel)

        If intYesNoPreCopy = vbOK Then
            intDoneCopying = MsgBox("Alright, so right now you should have a bunch of stuff in the clipboard." _
                                  & " If this is right, and you want to proceed with pasting and parsing the data, click Yes below", vbYesNo)
            If intDoneCopying = vbYes Then boolContinueWithPaste = True
            If intDoneCopying <> vbYes Then
                intRUSure = MsgBox("Cancel? Are you sure? You can always start again later, no data has been entered or deleted yet", vbYesNo)
                If intRUSure = vbYes Then End
                If intRUSure <> vbYes Then boolContinueWithPaste = False
            End If
        Else
            intRUSure = MsgBox("Cancel? Are you sure? You can always start again later, no data has been entered or deleted yet", vbYesNo)
            If intRUSure = vbYes Then End
            If intRUSure <> vbYes Then boolContinueWithPaste = False
        End If
    Wend

    'Paste data from clipboard, as ONLY text, into last sheet (called "PeopleSoft"), then fit columns
    Cells(1, 1).Activate
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    If WorksheetFunction.CountA(Range("B:B")) = 0 Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.Sheets("PeopleSoft").Delete
        On Error GoTo ErrHandlerCode
        intNumWorksheets = ThisWorkbook.Sheets.Count
        Sheets.Add(after:=Sheets(intNumWorksheets)).Name = "PeopleSoft"
        ActiveSheet.PasteSpecial Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
        Cells.Select
        Selection.ClearFormats
        Selection.ClearHyperlinks
        Application.DisplayAlerts = True
    End If
    Worksheets("PeopleSoft").Columns("A:Z").AutoFit


    'check that they didn't paste something weird
    Dim boolGoodPaste As Boolean


    boolGoodPaste = False



    For i = 1 To 50
        For j = 1 To 50
            If Cells(i, j) = "Status" Then
                boolGoodPaste = True
                If j > 4 Then Range(Cells(1, 1), Cells(150, j - 4)).EntireColumn.Delete
                i = 100
                j = 100
            End If
        Next j
    Next i

    If boolGoodPaste = False Then
        MsgBox ("Looks like you've pasted something else. This macro will now end; please read all instructions!")
        Sheets("Time Sheet Planner").Cells(1, 1).Activate
        End
    End If



    '
    ' Now, time to do the heavy lifting. First, to find the proper header column for the timesheet
    '

    Dim intHeaderRowNum As Integer
    Dim intSundayRowNum As Integer
    Dim arrTimesheetAllValues() As Variant
    Dim arrHeaderValues() As Variant
    Dim arrDayAndPunches() As Variant
    Dim varLastItemCheck As Variant
    Dim varLastDay As Variant
    Dim boolLastItem As Boolean
    Dim boolFoundLastCol As Boolean
    Dim intLastColNum As Integer

    intHeaderRowNum = WorksheetFunction.Match("Day", Range("B1:B100"), 0)
    'TO FIX: <TYPE MISMATCH> varLastDay = WorksheetFunction.Max(Range(("B" + intHeaderRowNum), ("B" + intHeaderRowNum + 14)))
    intSundayRowNum = WorksheetFunction.Match("Sun", Range("B1:B100"), 0)

    'check if it's really the last day, if not get the true last row number

    boolLastItem = False
    While boolLastItem = False
        varLastItemCheck = Cells(intSundayRowNum + 1, 2).Value
        If varLastItemCheck = "" Then
            boolLastItem = True
        Else
            intSundayRowNum = intSundayRowNum + 1
        End If
    Wend

    'get number of values (including header)

    Dim intNumDayLines As Integer

    intNumDayLines = intSundayRowNum - intHeaderRowNum

    i = 0    'rows, where 0 is the header
    j = 0    'columns, where 0 is the three-letter day

    'find last column
    While boolFoundLastCol <> True
        If Cells(intHeaderRowNum, j + 5).Value = "Date" Then
            boolFoundLastCol = True
            intLastColNum = j + 5
        Else
            j = j + 1
        End If
    Wend

    i = 0

    ReDim arrTimesheetAllValues(intNumDayLines, intLastColNum - 2)
    ReDim arrHeaderValues(intLastColNum + 2)
    Dim intInCounter As Integer
    Dim intOutCounter As Integer
    Dim strInCountString As String
    Dim strOutCountString As String


    intInCounter = 1
    intOutCounter = 2    ' "Out1" replaces Lunch manually later

    For i = 0 To intNumDayLines
        j = 0
        For j = 0 To intLastColNum - 2
            If Cells(intHeaderRowNum + i, 2 + j).Value = "In" Then
                strInCountString = "In" & intInCounter
                intInCounter = intInCounter + 1
                arrTimesheetAllValues(i, j) = strInCountString
            ElseIf Cells(intHeaderRowNum + i, 2 + j).Value = "Out" Then
                strOutCountString = "Out" & intOutCounter
                intOutCounter = intOutCounter + 1
                arrTimesheetAllValues(i, j) = strOutCountString
            Else
                arrTimesheetAllValues(i, j) = Cells(intHeaderRowNum + i, 2 + j).Value
            End If
            If i = 0 Then arrHeaderValues(j) = arrTimesheetAllValues(i, j)
        Next j
    Next i

    'OUTPUT CODE WENT HERE BEFORE FUNCTIONIZED

    'Convert all times from variant to time formatted, by adding space before AM/PM and changing format to HH:MM:SS format

    Dim intFirstTimePunchCol As Integer
    Dim intLastTimePunchCol As Integer
    Dim intNumColsOfTimes As Integer
    Dim rngTimesRange As Range
    Dim intLenOfTime As Integer
    Dim strTempTimeString As String
    Dim k As Integer

    intFirstTimePunchCol = findinarray("Status", arrHeaderValues) + 1
    intLastTimePunchCol = findinarray("Punch Total", arrHeaderValues) - 1
    intNumColsOfTimes = intLastTimePunchCol - intFirstTimePunchCol + 1

    ReDim arrDayAndPunches(intNumDayLines - 1, intNumColsOfTimes + 2)


    'replace all time values with spaces before AM/PM to format properly
    i = 1
    j = intFirstTimePunchCol - 1

    For i = 1 To intNumDayLines
        For j = (intFirstTimePunchCol - 1) To (intLastTimePunchCol - 1)
            intLenOfTime = Len(arrTimesheetAllValues(i, j))
            strTempTimeString = arrTimesheetAllValues(i, j)
            If Right(strTempTimeString, 2) = "AM" Then
                strTempTimeString = Left(arrTimesheetAllValues(i, j), intLenOfTime - 2) & " AM"
            ElseIf Right(strTempTimeString, 2) = "PM" Then
                strTempTimeString = Left(arrTimesheetAllValues(i, j), intLenOfTime - 2) & " PM"
            End If
            arrTimesheetAllValues(i, j) = strTempTimeString
        Next j
    Next i


    'clear the page (by deleting it and making it again)
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Sheets("PeopleSoft").Delete
    intNumWorksheets = ThisWorkbook.Sheets.Count
    Sheets.Add(after:=Sheets(intNumWorksheets)).Name = "PeopleSoft"
    Application.DisplayAlerts = True

    i = 0
    j = 0

    'OUTPUT BACK TO SHEET
    For i = 0 To intNumDayLines
        j = 0
        For j = 0 To intLastColNum - 2
            If (i = 0 And arrHeaderValues(j) = "Lunch") Then
                Cells(i + 1, j + 1) = "Out1"
            Else
                If (j >= intFirstTimePunchCol - 1 And j <= intLastTimePunchCol - 1 And i > 0) Then
                    Cells(i + 1, j + 1) = arrTimesheetAllValues(i, j)
                    Cells(i + 1, j + 1).NumberFormat = "h:mm:ss"
                Else
                    Cells(i + 1, j + 1) = arrTimesheetAllValues(i, j)
                End If
            End If
        Next j
    Next i

    Worksheets("PeopleSoft").Columns("A:Z").AutoFit

    Worksheets("PeopleSoft").Columns("A:Z").HorizontalAlignment = xlCenter


    'timestamp the sheet
    Cells(intNumDayLines + 8, 1) = "Pasted from PeopleSoft on " & Date & " at " & Time()
    Cells(intNumDayLines + 8, 1).HorizontalAlignment = xlLeft

    Set rngTimesRange = Range(Cells(2, intFirstTimePunchCol), Cells(intNumDayLines + 1, intLastTimePunchCol))


    'combine days on multiple lines logically, giving one string of <=6 punches (max 6 supported by main page)
    Dim intNumSubmitted As Integer
    Dim intBlankPunches As Integer
    Dim rngStatus As Range
    Dim boolAnyMultiLines As Boolean
    Dim intNumPunchLines As Integer
    Dim intNumPunchCols As Integer


    Set rngStatus = Range(Cells(2, intFirstTimePunchCol - 1), Cells(intNumDayLines + 1, intFirstTimePunchCol - 1))
    intNumPunchLines = rngStatus.Count

    If (intNumPunchLines > 7) Or (intInCounter = 3) Then
        Cells(1, intLastTimePunchCol).EntireColumn.Offset(0, 1).Insert
        Cells(1, intLastTimePunchCol).EntireColumn.Offset(0, 1).Insert
        Cells(1, intLastTimePunchCol).Offset(0, 1) = "In" & intInCounter
        Cells(1, intLastTimePunchCol).Offset(0, 2) = "Out" & intOutCounter
    End If

    'determine limits again, by re-creating header array and recounting everything

    ReDim arrHeaderValues(intLastColNum)

    For j = 0 To intLastColNum
        arrHeaderValues(j) = Cells(1, j + 1)
    Next j



    intFirstTimePunchCol = findinarray("Status", arrHeaderValues) + 1
    intLastTimePunchCol = findinarray("Punch Total", arrHeaderValues) - 1
    intNumColsOfTimes = intLastTimePunchCol - intFirstTimePunchCol + 1

    ReDim arrDayAndPunches(intNumDayLines - 1, intNumColsOfTimes + 2)




    'discover in first column if there are gaps, and which days have multiple lines
    i = 1
    j = intFirstTimePunchCol - 1

    'add days and punches (including blanks) to unique array (2D)
    For i = 1 To intNumDayLines
        arrDayAndPunches(i - 1, 0) = Cells(i + 1, 1).Value
    Next i


    For i = 1 To intNumDayLines
        For j = intFirstTimePunchCol To intLastTimePunchCol
            arrDayAndPunches(i - 1, j - 3) = Cells(i + 1, j).Value
        Next j
    Next i





    'find lines (row numbers) that have blank day values and need to be shifted up
    Dim arrDaysWithMultiPunchLines() As Variant
    Dim intDaysWithMultiPunchLines As Integer
    Dim arrDaysFixed(1 To 7) As String
    Dim strCurrentDay As String
    Dim intBlankLineCounter As Integer
    Dim intDayCounter As Integer
    Dim intCorrectMultiDays As Integer
    Dim boolMultiDaysExist As Boolean

    arrDaysFixed(1) = "Mon"
    arrDaysFixed(2) = "Tue"
    arrDaysFixed(3) = "Wed"
    arrDaysFixed(4) = "Thu"
    arrDaysFixed(5) = "Fri"
    arrDaysFixed(6) = "Sat"
    arrDaysFixed(7) = "Sun"

    boolMultiDaysExist = False
    intBlankLineCounter = 0
    intDayCounter = 0
    strCurrentDay = ""
    For i = 0 To intNumPunchLines - 1    'if the first column's day is blank, that day name (3 chars) is added to arrDaysWithMultiPunchLines to later check when moving punches
        If arrDayAndPunches(i, 0) = " " Or arrDayAndPunches(i, 0) = vbNullString Or arrDayAndPunches(i, 0) = "" Or arrDayAndPunches(i, 0) = Empty Then
            ReDim Preserve arrDaysWithMultiPunchLines(1 To intBlankLineCounter + 1)
            arrDaysWithMultiPunchLines(intBlankLineCounter + 1) = strCurrentDay
            intBlankLineCounter = intBlankLineCounter + 1
            boolMultiDaysExist = True
        Else
            intDayCounter = intDayCounter + 1
            strCurrentDay = arrDaysFixed(intDayCounter)
        End If
    Next i

    intDaysWithMultiPunchLines = UBound(arrDaysWithMultiPunchLines)

    'intCorrectMultiDays = MsgBox("The are " & intDaysWithMultiPunchLines & " lines with multiple punches. Is this correct?", vbOKCancel)
    '
    'If intCorrectMultiDays <> vbOK Then End

    '*****************************************************
    '       TASK ##
    '
    '       COMBINE ROWS OF PUNCHES, ENDING WITH 7
    '*****************************************************


    'for blank day rows, shift all non-blank cells up 1 and delete entire row

    If boolMultiDaysExist = True Then
        Dim intTargetRowToMovePunches As Integer
        Dim intLineOffset As Integer
        Dim intAbsLineOffset As Integer
        Dim strPreviousTargetDay As String
        Dim strCurrentTargetDay As String
        Dim intReusableCounter As Integer
        Dim intColumnOffset As Integer
        Dim boolPunchMoved As Boolean
        Dim intTargetColToMovePunches As Integer
        Dim intQuantityPunchesMoved As Integer
        Dim intQuantityPunchesToMove As Integer
        intQuantityPunchesToMove = 0
        intQuantityPunchesMoved = 0
        intLineOffset = -1

        strPreviousTargetDay = ""
        For i = 1 To intDaysWithMultiPunchLines
            intLineOffset = -1
            strCurrentTargetDay = arrDaysWithMultiPunchLines(i)
            If strCurrentTargetDay = strPreviousTargetDay Then intLineOffset = intLineOffset - 1
            intTargetRowToMovePunches = WorksheetFunction.Match(strCurrentTargetDay, Range("A1:A20"), 0)
            For j = intFirstTimePunchCol To intLastTimePunchCol
                If Cells(intTargetRowToMovePunches, j) = "" Or Cells(intTargetRowToMovePunches, j) = " " Then
                    Cells(intTargetRowToMovePunches, j) = "ISBLANK"
                End If
                If intLineOffset = -1 Then
                    If Cells(intTargetRowToMovePunches + 1, j) = "" Or Cells(intTargetRowToMovePunches + 1, j) = " " Then
                        Cells(intTargetRowToMovePunches + 1, j) = "ISBLANK"
                    Else
                        intQuantityPunchesToMove = intQuantityPunchesToMove + 1
                    End If
                ElseIf intLineOffset = -2 Then
                    If Cells(intTargetRowToMovePunches + 1, j) = "" Or Cells(intTargetRowToMovePunches + 1, j) = " " Or Cells(intTargetRowToMovePunches + 1, j) = "ISBLANK" Then
                        Cells(intTargetRowToMovePunches + 1, j) = "ISBLANK"
                    Else
                        intQuantityPunchesToMove = intQuantityPunchesToMove + 1
                    End If
                    If Cells(intTargetRowToMovePunches + 2, j) = "" Or Cells(intTargetRowToMovePunches + 2, j) = " " Then
                        Cells(intTargetRowToMovePunches + 2, j) = "ISBLANK"
                    Else
                        intQuantityPunchesToMove = intQuantityPunchesToMove + 1
                    End If
                End If
            Next j
            strPreviousTargetDay = strCurrentTargetDay
        Next i

        'intQuantityPunchesToMove = (intQuantityPunchesToMove - (intDaysWithMultiPunchLines * (intLastTimePunchCol - intFirstTimePunchCol)))

        'MsgBox (intQuantityPunchesToMove & " punches detected to move.")

        For i = 1 To intDaysWithMultiPunchLines
            intLineOffset = -1
            strCurrentTargetDay = arrDaysWithMultiPunchLines(i)
            If strCurrentTargetDay = strPreviousTargetDay Then intLineOffset = intLineOffset - 1
            intTargetRowToMovePunches = WorksheetFunction.Match(strCurrentTargetDay, Range("A1:A20"), 0)
            intReusableCounter = Abs(intLineOffset + 1)
            intAbsLineOffset = Abs(intLineOffset)
            For intTargetColToMovePunches = intFirstTimePunchCol To intLastTimePunchCol
                boolPunchMoved = False
                intColumnOffset = 0
                Cells(intTargetRowToMovePunches + intReusableCounter, intTargetColToMovePunches).Select
                If Cells(intTargetRowToMovePunches + intReusableCounter, intTargetColToMovePunches) <> "ISBLANK" Then  'find non-blank in offset row
                    Do While boolPunchMoved = False
                        If Cells(intTargetRowToMovePunches, intTargetColToMovePunches + intColumnOffset) = "ISBLANK" Then    'find somewhere to put it (the next ISBLANK above)
                            Cells(intTargetRowToMovePunches, intTargetColToMovePunches + intColumnOffset).Select
                            Cells(intTargetRowToMovePunches, intTargetColToMovePunches + intColumnOffset) = Cells(intTargetRowToMovePunches + intReusableCounter, intTargetColToMovePunches)
                            Cells(intTargetRowToMovePunches + intReusableCounter, intTargetColToMovePunches) = ""
                            boolPunchMoved = True
                            intQuantityPunchesMoved = intQuantityPunchesMoved + 1
                        Else
                            intColumnOffset = intColumnOffset + 1
                        End If
                    Loop
                    Cells(intTargetRowToMovePunches + intReusableCounter, intTargetColToMovePunches) = ""
                End If
            Next intTargetColToMovePunches
            strPreviousTargetDay = strCurrentTargetDay
        Next i
    End If

    'MsgBox (intQuantityPunchesMoved & " punches moved out of " & intQuantityPunchesToMove & " detected.")

    Dim intNumRowsDeleted As Integer
    Dim arrSingleArrayForSortedPunches() As Variant

    ReDim arrSingleArrayForSortedPunches(1)
    Sheets("PeopleSoft").Activate
    intReusableCounter = 1
    intNumRowsDeleted = 0
    i = WorksheetFunction.Match("Sun", Range("A1:A20"), 0)
    Do Until i = 1
        If Cells(i, 1) = " " Or Cells(i, 1) = vbNullString Or Cells(i, 1) = "Sun" Then    'Or Cells(i, 3) = "New" Then
            Rows(i).Delete
            intNumRowsDeleted = intNumRowsDeleted + 1
        End If
        i = i - 1
    Loop

    ReDim arrSingleArrayForSortedPunches(1 To ((intNumDayLines - intNumRowsDeleted) * intNumColsOfTimes))
    Dim intFinalNumDayRows As Integer

    intFinalNumDayRows = intNumDayLines - intNumRowsDeleted + 1


    For i = 2 To intFinalNumDayRows
        For j = intFirstTimePunchCol To intLastTimePunchCol
            If Cells(i, 1) = "" Or Cells(i, 1) = " " Then
                'do nothing
            Else
                If Cells(i, j) = " " Or Cells(i, j) = "ISBLANK" Or Cells(i, j) = Empty Then
                    Cells(i, j) = ""
                    arrSingleArrayForSortedPunches(intReusableCounter) = ""
                Else
                    arrSingleArrayForSortedPunches(intReusableCounter) = Cells(i, j).Value
                End If
                intReusableCounter = intReusableCounter + 1
            End If
        Next j
    Next i



    'pull any punches from first sheet, and compare if necessary. Prompt for conflicts
    'first, back up Sheet(1) just in case anything goes weird

    Dim intFirstSheetTargetRowOne As Integer
    Dim strFirstSheetName As String
    Dim arrSingleArrayForExistingPunches() As Variant

    Sheets(1).Activate
    strFirstSheetName = Sheets(1).Name
    intFirstSheetTargetRowOne = WorksheetFunction.Match("Monday", Range("A1:A16"), 0)

    ReDim arrSingleArrayForExistingPunches(1 To UBound(arrSingleArrayForSortedPunches))

    intReusableCounter = 1
    For i = intFirstSheetTargetRowOne To (intFirstSheetTargetRowOne + (intFinalNumDayRows * 2)) Step 2
        For j = 2 To 7
            Cells(i, j).Activate
            If Cells(i, j) <> Empty Then
                arrSingleArrayForExistingPunches(intReusableCounter) = Cells(i, j).Value
            Else
                arrSingleArrayForExistingPunches(intReusableCounter) = ""
            End If
            intReusableCounter = intReusableCounter + 1
        Next j
    Next i


    'COMPARE AND REPORT CONFLICTS, ROUNDING TO RULES FIRST
    Dim dblTimeExisting#
    Dim dblTimeSorted#
    Dim dblHoursExisting#
    Dim dblHoursExistingFloor#
    Dim dblHoursSorted#
    Dim dblHoursSortedFloor#
    Dim dblMinExisting#
    Dim dblMinExistingFloor#
    Dim dblMinSorted#
    Dim dblMinSortedFloor#
    Dim dblSecExistingRound#
    Dim dblSecSortedRound#

    Dim dblRoundedDecimalTimeExisting#
    Dim dblRoundedDecimalTimeSorted#

    Dim arrSortedCalculatedPunches As Variant
    Dim arrExistingCalculatedPunches As Variant

    Dim dblIntTest#



    intAllArrayLengths = UBound(arrSingleArrayForSortedPunches)

    ReDim arrDoConflictsExist(1 To intAllArrayLengths)
    ReDim arrSortedCalculatedPunches(1 To (intAllArrayLengths), 1 To 5)
    ReDim arrExistingCalculatedPunches(1 To (intAllArrayLengths), 1 To 5)
    ' above redim:
    '   (i,1): existing times for array
    '   (i,2): hours integer [floor (i,1)*24]
    '   (i,3): minutes integer [floor (i,2)*60]
    '   (i,4): seconds integer [round (i,3)*60]
    '   (i,5): rounded decimal time, using PeopleSoft rounding rules

    For i = 1 To intAllArrayLengths
        arrSortedCalculatedPunches(i, 1) = arrSingleArrayForSortedPunches(i)
        arrExistingCalculatedPunches(i, 1) = arrSingleArrayForExistingPunches(i)
    Next i

    For i = 1 To intAllArrayLengths
        If arrExistingCalculatedPunches(i, 1) = "" Then
            'foo
        Else
            dblTimeExisting = arrExistingCalculatedPunches(i, 1)
            dblHoursExisting = dblTimeExisting * 24
            dblHoursExistingFloor = Int(dblHoursExisting)
            dblMinExisting = (dblHoursExisting - dblHoursExistingFloor) * 60
            dblIntTest = dblMinExisting - Int(dblMinExisting)
            dblIntTest = Round(dblIntTest + 0.000000001, 4)
            If dblIntTest <> 1 Then
                dblMinExistingFloor = Int(dblMinExisting)
            Else
                dblMinExistingFloor = dblMinExisting
            End If
            dblSecExistingRound = Round(((dblMinExisting - dblMinExistingFloor) * 60) + 0.0000001, 0)
            If dblSecExistingRound = 60 Then dblSecExistingRound = 0
            arrExistingCalculatedPunches(i, 2) = dblHoursExistingFloor
            arrExistingCalculatedPunches(i, 3) = dblMinExistingFloor
            arrExistingCalculatedPunches(i, 4) = dblSecExistingRound
        End If
        '
        '
        If arrSortedCalculatedPunches(i, 1) = "" Then
            'foo
        Else
            dblTimeSorted = arrSortedCalculatedPunches(i, 1)
            dblHoursSorted = dblTimeSorted * 24
            dblHoursSortedFloor = Int(dblHoursSorted)
            dblMinSorted = (dblHoursSorted - dblHoursSortedFloor) * 60
            dblIntTest = dblMinSorted - Int(dblMinSorted)
            If dblIntTest <> 1 Then
                dblMinSortedFloor = Int(dblMinSorted)
            Else
                dblMinSortedFloor = dblMinSorted
            End If
            dblSecSortedRound = Round(((dblMinSorted - dblMinSortedFloor) * 60) + 0.0000001, 0)
            If dblSecSortedRound = 60 Then dblSecSortedRound = 0

            arrSortedCalculatedPunches(i, 2) = dblHoursSortedFloor
            arrSortedCalculatedPunches(i, 3) = dblMinSortedFloor
            arrSortedCalculatedPunches(i, 4) = dblSecSortedRound
        End If

    Next i

    ' then use if/then to round each of the times, then store again as decimal at (i,5)
    ' ROUNDING CASES (decimal minutes):
    '  0   <= x < 7.5   = 00
    '  7.5 <= x < 22.5  = 15
    '  22.5<= x < 37.5  = 30
    '  37.5<= x < 52.5  = 45
    '  52.5<= x <=60    = 00, next hour

    Dim x#
    Dim intHrs As Integer
    Dim intMin As Integer

    'round existing, inserting rounded decimal time into array(i,5) as #
    For i = 1 To intAllArrayLengths
        intMin = arrExistingCalculatedPunches(i, 3)
        x = intMin + (arrExistingCalculatedPunches(i, 4) / 60)
        intHrs = arrExistingCalculatedPunches(i, 2)
        If arrExistingCalculatedPunches(i, 1) = "" Then
            'foo
        Else
            If (x >= 0 And x < 7.5) Then    '00
                arrExistingCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 0)
            ElseIf (x >= 7.5 And x < 22.5) Then    '15
                arrExistingCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 15)
            ElseIf (x >= 22.5 And x < 37.5) Then    '30
                arrExistingCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 30)
            ElseIf (x >= 37.5 And x < 52.5) Then    '45
                arrExistingCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 45)
            ElseIf (x >= 52.5 And x <= 60) Then    ' 00, next hour
                arrExistingCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs + 1, 0)
            End If
        End If

        'round PeopleSoft (sorted)
        'repeat above code

        intMin = arrSortedCalculatedPunches(i, 3)
        x = intMin + (arrSortedCalculatedPunches(i, 4) / 60)
        intHrs = arrSortedCalculatedPunches(i, 2)
        If arrSortedCalculatedPunches(i, 1) = "" Then
            'foo
        Else
            If (x >= 0 And x < 7.5) Then    '00
                arrSortedCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 0)
            ElseIf (x >= 7.5 And x < 22.5) Then    '15
                arrSortedCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 15)
            ElseIf (x >= 22.5 And x < 37.5) Then    '30
                arrSortedCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 30)
            ElseIf (x >= 37.5 And x < 52.5) Then    '45
                arrSortedCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs, 45)
            ElseIf (x >= 52.5 And x <= 60) Then    ' 00, next hour
                arrSortedCalculatedPunches(i, 5) = MakeTimeDecimal(intHrs + 1, 0)
            End If
        End If

    Next i


    Dim dblDiff#
    Dim dbltimeDiff#
    Dim intNumEmptyConflicts As Integer
    Dim intNumFalseConflicts As Integer
    Dim tmpSorted#
    Dim tmpExisting#
    Dim strExistingString As String
    Dim strSortedString As String

    i = 1
    arrDoConflictsExist(i) = False
    intNumEmptyConflicts = 0
    intNumFalseConflicts = 0
    intNumTrueConflicts = 0

    'now, see if conflicts exist. If so, store bool TRUE / FALSE in arrDoConflictsExist (default FALSE)
    For i = 1 To intAllArrayLengths
        tmpSorted = arrSortedCalculatedPunches(i, 5)
        tmpExisting = arrExistingCalculatedPunches(i, 5)
        If (arrSortedCalculatedPunches(i, 5) = Empty And arrExistingCalculatedPunches(i, 5) = Empty) Then
            arrDoConflictsExist(i) = Empty
            intNumEmptyConflicts = intNumEmptyConflicts + 1
        Else
            strExistingString = MakeTimeString(tmpExisting)
            strSortedString = MakeTimeString(tmpSorted)
            If (strExistingString = strSortedString) Or strExistingString = "No Punch" Then
                arrDoConflictsExist(i) = False
                intNumFalseConflicts = intNumFalseConflicts + 1
            Else
                arrDoConflictsExist(i) = True
                intNumTrueConflicts = intNumTrueConflicts + 1
                'MsgBox ("Time Existing: " & strExistingString & vbCrLf & "Time Sorted: " & strSortedString)
            End If
        End If
    Next i

    i = 2


    ReDim arrSortedForm(1 To intAllArrayLengths)
    ReDim arrExistingForm(1 To intAllArrayLengths)
    ReDim arrFinalPunchesToUse(1 To (intAllArrayLengths + 10), 1 To 2)

    For i = 1 To UBound(arrFinalPunchesToUse)
        arrFinalPunchesToUse(i, 2) = "P"
    Next i


    'arrSortedForm = arrSortedCalculatedPunches
    'arrExistingForm = arrExistingCalculatedPunches

    For i = 1 To intAllArrayLengths
        tmpSorted = arrSortedCalculatedPunches(i, 5)
        tmpExisting = arrExistingCalculatedPunches(i, 5)
        If (tmpSorted = Empty) And (tmpExisting = Empty) Then
            arrSortedForm(i) = Empty
            arrExistingForm(i) = Empty
        Else
            arrSortedForm(i) = MakeTimeString(tmpSorted)
            arrExistingForm(i) = MakeTimeString(tmpExisting)
        End If
    Next i


    i = 1

    frmPunchConflictReview.Show
    Unload frmPunchConflictReview


    'copy single-line punches from PeopleSoft to main page (values only, using:
    'ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True

    intReusableCounter = 1
    Dim rngWorkingCell As Range

    Sheets("Time Sheet Planner").Activate

    For i = intFirstSheetTargetRowOne To (intFirstSheetTargetRowOne + ((intFinalNumDayRows - 1) * 2)) Step 2
        For j = 2 To 7
            Cells(i, j).Activate
            Cells(i, j).Value = 0
            With Cells(i, j).Interior    'clear fill
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            If arrFinalPunchesToUse(intReusableCounter, 1) = "No Punch" Then
                Cells(i, j).Value = ""
            Else
                Cells(i, j).Value = arrFinalPunchesToUse(intReusableCounter, 1)
            End If
            Cells(i, j).NumberFormat = "[$-F400]h:mm:ss AM/PM"
            If arrFinalPunchesToUse(intReusableCounter, 2) = "P" And Cells(i, j).Value <> "" Then
                With Cells(i, j).Interior    'shade red
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            End If
            intReusableCounter = intReusableCounter + 1
            'Set rngWorkingCell = Range(Cells(i, j))
            On Error Resume Next
            ActiveCell.ClearComments
            ActiveCell.AddComment
            ActiveCell.Comment.Text CStr(Cells(i, j).Value), 1, 1
            'Cells(i, j).Comment.Text rngWorkingCell.Value
            'rngWorkingCell.Comment.Text rngWorkingCell.Value
            On Error GoTo 0
            'ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
        Next j
    Next i

    Application.DisplayCommentIndicator = xlNoIndicator


    'Dim intMonCount As Integer
    'Dim intTueCount As Integer
    'Dim intWedCount As Integer
    'Dim intThuCount As Integer
    'Dim intFriCount As Integer
    'Dim intSatCount As Integer
    '
    'intMonCount = WorksheetFunction.CountA(Range("B3:G3"))
    'intTueCount = WorksheetFunction.CountA(Range("B5:G5"))
    'intWedCount = WorksheetFunction.CountA(Range("B7:G7"))
    'intThuCount = WorksheetFunction.CountA(Range("B9:G9"))
    'intFriCount = WorksheetFunction.CountA(Range("B11:G11"))
    'intSatCount = WorksheetFunction.CountA(Range("B13:G13"))
    '
    'Call FixOffsetPunches(5, intTueCount)

    With Cells(17, 2)
        .Value = "Note: Red shaded cells denote times imported from PeopleSoft"
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With


    MsgBox ("Complete! If something went weird, there's a backup of your original first sheet (see tabs at bottom)." _
          & vbCrLf & vbCrLf & "If you don't need it, you can delete that sheet at any time.")


    Dim boolRestoreBackup

    boolRestoreBackup = MsgBox("Delete the backup now and use populated values?" _
                             & vbCrLf & vbCrLf & "Yes: Keep changes to main sheet and delete backup." _
                             & vbCrLf & "No: Restore my old stuff (undoes changes made by running this program)." _
                             & vbCrLf & "Cancel: Do nothing (keep both sheets)." _
                             & vbCrLf & vbCrLf & "(WARNING! Any action is not reversible.)", vbYesNoCancel)

    Select Case boolRestoreBackup
        Case vbYes
            'Delete backup sheet
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Backup of Time Sheet Planner").Delete
            Application.DisplayAlerts = True
        Case vbNo
            'Restore backup sheet (pre-macro state)
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets("Time Sheet Planner").Delete
            Sheets("Backup of Time Sheet Planner").Name = "Time Sheet Planner"
            Application.DisplayAlerts = True
        Case vbCancel
            'foo
    End Select



    ActiveSheet.Range("K15").Activate

    'ActiveWorkbook.Save

    Application.EnableEvents = True

    Sheets("Time Sheet Planner").Activate
    ActiveSheet.Cells(1, 1).Select



    '   **********************************************************************
    '
    '           ERROR REPORTING CODE
    '
    '   **********************************************************************

    Err.Number = 0
ErrHandlerCode:
    If Err.Number <> 0 Then
        MsgBox ("Woops! I've encountered an error I didn't plan for." & vbCrLf & vbCrLf & "Please report this error to the developer:" _
              & vbCrLf & vbCrLf & "Error # " & Str(Err.Number) & ": " & Err.Description & vbCrLf & vbCrLf & vbCrLf & "Running the program again will probably make the error go away.")
    End If


    Call MaintenanceAndRepair

End Sub



'finds a variable inside an array and returns the relative position as an integer
Public Function findinarray(searchvalue As Variant, arr As Variant) As Integer
    On Error Resume Next
    findinarray = Application.Match(searchvalue, arr, 0)
    On Error GoTo 0
End Function

'takes hour and minute integers, after being rounded using PeopleSoft rules, and makes them decimals again
Public Function MakeTimeDecimal(roundedhours As Integer, roundedminutes As Double) As Double
    ' doesn't deal with seconds, assume :00
    MakeTimeDecimal = (roundedhours + (roundedminutes / 60)) / 24
End Function

'takes decimal times, after being rounded using PeopleSoft rules, and makes them formatted strings for form display
Public Function MakeTimeString(decimalhours As Double) As String
    Dim strAMPM As String
    Dim strHrs As String
    Dim strMin As String
    Dim strTime As String
    Dim roundhours As Integer
    Dim roundmin As Integer
    ' doesn't deal with seconds, assume :00
    strAMPM = ""
    roundhours = Round(Int(decimalhours * 24.000000000001) + 0.000000001, 4)
    roundmin = Round((((decimalhours * 24) - Int(decimalhours * 24)) * 60) + 0.00000001, 4)
    If roundhours >= 12 Then strAMPM = " PM" Else strAMPM = " AM"
    If roundhours > 12 Then roundhours = roundhours - 12
    If roundhours < 10 Then strHrs = ("0" & roundhours) Else strHrs = roundhours
    If (roundmin = 0 Or roundmin = 60) Then strMin = "00" Else strMin = roundmin
    strTime = strHrs & ":" & strMin & strAMPM
    If (strTime = "00:00 AM") Or (strTime = "00:00 PM") Then MakeTimeString = "No Punch" Else MakeTimeString = strTime
End Function


'given a punch's location within the array, find its day and in/out status
Public Function FindDayOfPunch(intPosInArray As Integer) As String
    Dim strDayName As String
    Select Case intPosInArray
        Case 1 To 6
            strDayName = "Monday"
        Case 7 To 12
            strDayName = "Tuesday"
        Case 13 To 18
            strDayName = "Wednesday"
        Case 19 To 24
            strDayName = "Thursday"
        Case 25 To 30
            strDayName = "Friday"
        Case 31 To 36
            strDayName = "Saturday"
        Case 37 To 42
            strDayName = "Sunday"
        Case Else
            strDayName = "<Day Unknown>"
    End Select

    FindDayOfPunch = strDayName

End Function



'given a punch's location within the array, find it's day and in/out status
Public Function FindStatusOfPunch(intPosInArray As Integer) As String
    Dim strInOutName As String
    Select Case intPosInArray
        Case 1, 7, 13, 19, 25, 31, 37
            strInOutName = "In1"
        Case 2, 8, 14, 20, 26, 32, 38
            strInOutName = "Out1"
        Case 3, 9, 15, 21, 27, 33, 39
            strInOutName = "In2"
        Case 4, 10, 16, 22, 28, 34, 40
            strInOutName = "Out2"
        Case 5, 11, 17, 23, 29, 35, 41
            strInOutName = "In3"
        Case 6, 12, 18, 24, 30, 36, 42
            strInOutName = "Out3"
        Case Else
            strInOutName = "<In/Out Unknown>"
    End Select

    FindStatusOfPunch = strInOutName

End Function


Public Function FixOffsetPunches(intDayRow As Integer, intCountA As Integer)
    Dim vbContinueFix

    vbContinueFix = MsgBox("Looks like you're missing a punch for " & Cells(intDayRow, 1).Value & ". Would you like to fix this and add one?" & vbCrLf & vbCrLf _
                         & "This can be useful if you didn't punch in at the start of a day.", vbYesNo)
    If vbContinueFix <> vbYes Then End

    For i = 1 To intCountA

    Next i



End Function


