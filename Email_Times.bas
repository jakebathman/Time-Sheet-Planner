Attribute VB_Name = "Email_Times"
Option Explicit
Public arrPeopleAndEmails() As String
Public arrPeople() As String
Public strEmail As String
Public strName As String
Public arrDaysSelected() As String
Public boolPreviousWeek As Boolean
Public c As Integer
Public strWeekString As String
Public strDate As Date
Public strMonth As String
Public strDayofMonth As String
Public strSheetName As String
Public boolRedPunches As Boolean
Public arrCheckBoxStates(1 To 7) As Boolean
Public intOtherEmailsFirstRow As Integer
Public boolDone As Boolean

Public Sub EmailTimes()

Call MaintenanceAndRepair

On Error GoTo ErrHandlerCode


Dim i As Integer
Dim j As Integer
Dim intMondayRow As Integer
Dim arrTimes() As String
Dim intCounter As Integer
Dim OutApp As Object
Dim OutMail As Object
Dim strMessage As String
Dim strMessageFormatted As String
Dim arrHeaders() As String
Dim strSigPath As String
Dim strSigString As String
Dim boolNoPunchesRunImport
Erase arrPeople()
Erase arrPeopleAndEmails()

ReDim arrPeopleAndEmails(1 To 6, 1 To 2)
ReDim arrPeople(1 To 6, 1 To 2)

'Unload frmPickDaysToEmail
'Unload frmSelectPersonToEmail
'Unload frmWorking

'check to see if there are times available to email
boolNoPunchesRunImport = vbNo
For i = 3 To 9
    For j = 2 To 7
        If Cells(i, j).Value = "" Or Cells(i, j).Value = Empty Then
            With Cells(i, j).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            i = 100: j = 100
            boolNoPunchesRunImport = vbCancel
        End If
    Next j
Next i

If i < 50 Then boolNoPunchesRunImport = MsgBox("Looks like you don't have any punches to email." & vbCrLf & vbCrLf & "You should add some first.", vbOKOnly)


If boolNoPunchesRunImport = vbYes Then
    Call PeopeSoftImport
    End ' prevents the rest of this sub from running after import
End If

arrPeople(1, 1) = "Jake B."
arrPeople(2, 1) = "Oscar M."
arrPeople(3, 1) = "Carol S."
arrPeople(4, 1) = "Kelley S."
arrPeople(5, 1) = "Other A"
arrPeople(6, 1) = "Other B"
If Sheets("User Preferences").Cells(13, 2).Value <> "" Then arrPeople(5, 1) = Sheets("User Preferences").Cells(13, 2).Value
If Sheets("User Preferences").Cells(14, 2).Value <> "" Then arrPeople(6, 1) = Sheets("User Preferences").Cells(14, 2).Value

arrPeopleAndEmails(1, 1) = "Jake B."
arrPeopleAndEmails(2, 1) = "Oscar M."
arrPeopleAndEmails(3, 1) = "Carol S."
arrPeopleAndEmails(4, 1) = "Kelley S."
arrPeopleAndEmails(5, 1) = "Other A"
arrPeopleAndEmails(6, 1) = "Other B"
If Sheets("User Preferences").Cells(13, 2).Value <> "" Then arrPeopleAndEmails(5, 1) = Sheets("User Preferences").Cells(13, 2).Value
If Sheets("User Preferences").Cells(14, 2).Value <> "" Then arrPeopleAndEmails(6, 1) = Sheets("User Preferences").Cells(14, 2).Value

arrPeopleAndEmails(1, 2) = "jbathman@co.collin.tx.us"
arrPeopleAndEmails(2, 2) = "omartinez@co.collin.tx.us"
arrPeopleAndEmails(3, 2) = "cstrickland@co.collin.tx.us"
arrPeopleAndEmails(4, 2) = "kstone@co.collin.tx.us"
arrPeopleAndEmails(5, 2) = ""
arrPeopleAndEmails(6, 2) = ""
If Sheets("User Preferences").Cells(13, 3).Value <> "" Then arrPeopleAndEmails(5, 2) = Sheets("User Preferences").Cells(13, 3).Value
If Sheets("User Preferences").Cells(14, 3).Value <> "" Then arrPeopleAndEmails(6, 2) = Sheets("User Preferences").Cells(14, 3).Value

ReDim arrHeaders(1 To 6)
arrHeaders(1) = "In"
arrHeaders(2) = "Out Lunch"
arrHeaders(3) = "In Lunch"
arrHeaders(4) = "Out"
arrHeaders(5) = "In 2"
arrHeaders(6) = "Out 2"

boolPreviousWeek = True
boolRedPunches = False
'Load frmWorking
intOtherEmailsFirstRow = 13

On Error Resume Next
intOtherEmailsFirstRow = WorksheetFunction.Match("Other emails", Sheets("User Preferences").Range("A:A"), 0) + 1
On Error GoTo ErrHandlerCode

boolDone = False

While boolDone = False
    frmSelectPersonToEmail.Show
    On Error Resume Next
    If boolDone = False Then
        If MsgBox("Really cancel?" & vbCrLf & vbCrLf & "Pressing ""No"" will bring back the dialog box.", vbYesNo) = vbYes Then Unload frmSelectPersonToEmail: End
    End If
    On Error GoTo ErrHandlerCode
Wend

If MsgBox("Going to email " & strEmail & vbCrLf & vbCrLf & "Continue?", vbYesNo) <> vbYes Then End

Sheets("Time Sheet Planner").Activate
intMondayRow = WorksheetFunction.Match("Monday", Range("A1:A15"), 0)

frmPickDaysToEmail.Show

ReDim arrTimes(1 To UBound(arrDaysSelected), 1 To 7)
intCounter = 1
strMessage = vbTab
For i = 2 To 7
    strMessage = strMessage & vbTab & vbTab & Cells(intMondayRow - 1, i).Value
Next i

For i = intMondayRow To (intMondayRow + 6)
    If findinarray(Cells(i, 1).Value, arrDaysSelected) <> 0 Then
        strMessage = strMessage & vbCrLf & Left(Cells(i, 1).Value, 3) & vbCrLf & vbTab
        arrTimes(intCounter, 1) = Cells(i, 1).Value
        For j = 2 To 7
            Cells(i, j).Activate
            arrTimes(intCounter, j) = MakeTimeString(Cells(i, j).Value)
            If arrTimes(intCounter, j) <> "No Punch" Then strMessage = strMessage & vbTab & "   " & MakeTimeString(Cells(i, j).Value) Else strMessage = strMessage & vbTab & vbTab
        Next j
        intCounter = intCounter + 1
    End If
Next i


If boolPreviousWeek = True Then c = 7 Else c = 0
strDate = SetDateStrings(Date)
strMonth = MonthName(Month(strDate))
strDayofMonth = Day(strDate)

Worksheets.Add after:=Sheets(ActiveWorkbook.Sheets.Count)
strSheetName = ActiveSheet.Name



For i = 1 To 6
    Cells(1, i + 1) = arrHeaders(i)
Next i

For i = 1 To intCounter - 1
    For j = 1 To 7
        If arrTimes(i, j) <> "No Punch" Then Cells(i + 1, j) = arrTimes(i, j) Else Cells(i + 1, j) = ""
    Next j
Next i

Sheets(strSheetName).Activate

If WorksheetFunction.CountA(Range("F2:G" & UBound(arrDaysSelected) + 1)) = 0 Then Range("F1:G1").Clear



If MsgBox("Would you like to differentiate any punches? This can be helpful" _
    & " if you want to ask a supervisor to add a particular punch for you." & vbCrLf & vbCrLf _
    & "Flagged punches will show up red.", vbYesNo) = vbYes Then frmFlagPunchesForEmail.Show
    

Application.EnableEvents = False

'frmWorking.Show vbModeless
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)


Range("A1:G1").Font.Bold = True
Range("A1:A" & UBound(arrDaysSelected) + 1).Font.Bold = True
Range("A1:G" & UBound(arrDaysSelected) + 1).HorizontalAlignment = xlCenter

'get signature
strSigPath = "C:\Documents and Settings\" & Environ("username") & "\Application Data\Microsoft\Signatures\"

If Dir(strSigPath) <> "" Then
    strSigString = FindSigFile(strSigPath)
Else
    strSigString = ""
End If


Dim strSigFont As String
Dim intSigFontSize As Integer

strSigFont = "Arial"
intSigFontSize = 10

On Error Resume Next
strSigFont = GetSignatureFont(strSigString)
intSigFontSize = GetSignatureFontSize(strSigString)
On Error GoTo ErrHandlerCode

Sheets(strSheetName).Activate
Cells.Select
With Selection.Font
    .Name = strSigFont
    .Size = intSigFontSize
End With

Dim strDateAsString As String
Dim OutlookStartTime, OutlookEndTime
strDateAsString = CStr(strDate)

'Unload frmWorking

On Error Resume Next
With OutMail
    '.display
    .To = strEmail
    .CC = ""
    .BCC = ""
    .Subject = "My times for the week of " & Mid(strDateAsString, 1, Len(strDateAsString) - 5)
    OutlookStartTime = Timer
    .htmlbody = strName & "," & "<br><br>" & "Here are my times for last week (starting Monday, " & CStr(Mid(strDateAsString, 1, Len(strDateAsString) - 5)) & "). "
    If boolRedPunches Then .htmlbody = .htmlbody & " Some times were marked with <font color=""FF0000"">red</font color> to differentiate them as needing to be added/changed in PeopleSoft."
    .htmlbody = .htmlbody & RangetoHTML(Sheets(strSheetName).Range("A1:G" & UBound(arrDaysSelected) + 1))  'OR strMessage
    .htmlbody = .htmlbody & "<br>" & ExtractTextInsideSpan(GetBoiler(strSigString, strSigFont, intSigFontSize))
    'You can add a file like this
    '.Attachments.Add ("C:\test.txt")
End With

OutlookEndTime = Timer

Application.ScreenUpdating = True
Application.ActiveWindow.Activate
Application.ActiveWorkbook.Sheets("Time Sheet Planner").Cells(intMondayRow, 2).Activate
If MsgBox("Email created. Send now? (Click No to edit before sending)", vbYesNo) = vbYes Then OutMail.send Else OutMail.display

On Error GoTo ErrHandlerCode

Set OutMail = Nothing
Set OutApp = Nothing

Application.DisplayAlerts = False
Sheets(strSheetName).Delete
Application.DisplayAlerts = True

Unload frmPickDaysToEmail
Unload frmSelectPersonToEmail
'Unload frmWorking

Sheets("Time Sheet Planner").Activate
Cells(1, 1).Activate

Application.EnableEvents = True
Application.ScreenUpdating = True

Dim timediff

timediff = OutlookEndTime - OutlookStartTime
Debug.Print timediff
If timediff > 5 Then
    If MsgBox("Looks like Outlook is showing you pesky security prompts." & vbCrLf & vbCrLf _
        & "Would you like to view easy instructions to make that go away? (recommended!)", vbYesNo) = vbYes Then
            MsgBox ("Alrighty. Unfortunately, there are a few options and none of them are perfect. You can do one of three things:" & vbCrLf & vbCrLf _
                & "1. Each time you're prompted by Outlook, simply click Allow (for one-time allowance) or Allow for a certain amount of time." & vbCrLf _
                & "2. Lower Outlook macro trust security setting (in Outlook: File >> Options >> Trust Center >> Trust Center Settings >> Macro Settings >> ""Enable all macros..."") " & vbCrLf _
                & "3. Install Microsoft Security Essentials and enable real-time protection (free program, easily found using Google).")
    End If
End If




Err.Number = 0
ErrHandlerCode:
    If Err.Number <> 0 Then
        MsgBox ("Woops! I've encountered an error I didn't plan for." & vbCrLf & vbCrLf & "Please report this error to the developer:" _
            & vbCrLf & vbCrLf & "Error # " & Str(Err.Number) & ": " & Err.Description & vbCrLf & vbCrLf & vbCrLf & "Running the program again will probably make the error go away.")
    End If


Call MaintenanceAndRepair

End Sub




Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2010
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteAll
        '.Cells(1).PasteSpecial xlPasteValues, , False, False
        '.Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).openastextstream(1, -2)
    RangetoHTML = ts.readall
    'Shell "notepad.exe " & TempFile
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function





Public Function FindSigFile(folderspec As String)

Dim Fs, f, f1, fc, tp, nm, s, temp As Date, myname As String
Set Fs = CreateObject("Scripting.FileSystemObject")
'folderspec = "c:\backup"
Set f = Fs.GetFolder(folderspec) ': Set fc = f.subfolders
'For Each f1 In fc
'   If f1.datecreated > temp Then temp = f1.datecreated: myname = f1.Path
'Next
Set fc = f.Files ': set f = fs.GetFolder(myname)
For Each f1 In fc
    nm = f1.Name
    tp = f1.Type
    If InStr(1, tp, "htm", vbTextCompare) Then
        If f1.Size > temp Then temp = f1.Size: myname = f1.Name
    End If
Next

FindSigFile = folderspec & myname

End Function


Function GetBoiler(ByVal sFile As String, sfont As String, ifont As Integer) As String


On Error Resume Next

    Dim fsob As Object
    Dim tsob As Object
    Dim strBoilerAll As String
    Dim strBoilerHeader As String
    Dim boolStillFindingFonts As Boolean
    Dim intLastFontPos As Integer
    Dim i As Integer
    Set fsob = CreateObject("Scripting.FileSystemObject")
    Set tsob = fsob.GetFile(sFile).openastextstream(1, -2)
    strBoilerAll = tsob.readall
    
'    i = Len(strBoilerAll)
'
'    strBoilerHeader = "<html><head><style> Table.MsoNormalTable {mso-style-name:""Table Normal""; mso-tstyle-rowband-size:0; mso-tstyle-colband-size:0; mso-style-noshow:yes;" _
'        & "mso-style-priority:99; mso-style-parent:""""; mso-padding-alt:0in 5.4pt 0in 5.4pt; mso-para-margin:0in; mso-para-margin-bottom:.0001pt; mso-pagination:widow-orphan;" _
'        & "font-size:11.0pt; font-family:""Calibri"",""serif""; mso-ascii-font-family:Calibri; mso-ascii-theme-font:minor-latin; mso-hansi-font-family:Calibri;" _
'        & "mso-hansi-theme-font:minor-latin;}</style></head>"
'    boolStillFindingFonts = True
'    intLastFontPos = 1
'    While boolStillFindingFonts
'        If InStr(intLastFontPos + 1, strBoilerHeader, "font-family:", vbBinaryCompare) <> 0 Then
'            intLastFontPos = InStr(intLastFontPos + 1, strBoilerHeader, "font-family:", vbBinaryCompare) + Len("font-family:")
'            strBoilerHeader = Mid(strBoilerHeader, 1, intLastFontPos) & """" & sfont & """" & Mid(strBoilerHeader, InStr(intLastFontPos + 1, strBoilerHeader, ";", vbBinaryCompare))
'        Else
'            boolStillFindingFonts = False
'        End If
'    Wend
    
    GetBoiler = Mid(strBoilerAll, InStr(1, strBoilerAll, "<body", vbBinaryCompare))
    tsob.Close
    Set tsob = Nothing
    Set fsob = Nothing
End Function


Public Function SetDateStrings(dt)

Dim strTempDate As Date
Dim i As Integer
strTempDate = dt - c
strWeekString = Weekday(strTempDate, vbMonday)
i = 1
While strWeekString > 1
    strTempDate = strTempDate - 1
    strWeekString = Weekday(strTempDate, vbMonday)
    i = i + 1
Wend

SetDateStrings = strTempDate

End Function


Function GetSignatureFont(PathToSig As String) As String

On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Dim strSigTextAll As String
    Dim strFontName As String
    Dim intLocInFile As Integer
    Dim intLocOfEnd As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(PathToSig).openastextstream(1, -2)
    strSigTextAll = ts.readall
    
    
    
    If InStr(InStr(1, strSigTextAll, "<body"), strSigTextAll, ";font-family:", vbBinaryCompare) <> 0 Then
        intLocInFile = InStr(InStr(1, strSigTextAll, "<body"), strSigTextAll, ";font-family:", vbBinaryCompare) + Len(";font-family:") + 1
        intLocOfEnd = InStr(intLocInFile, strSigTextAll, """", vbBinaryCompare)
        strFontName = Mid(strSigTextAll, intLocInFile, intLocOfEnd - intLocInFile)
    End If
    ts.Close
    GetSignatureFont = strFontName
    Set ts = Nothing
    Set fso = Nothing
End Function


Function GetSignatureFontSize(PathToSig As String) As String

On Error Resume Next

    Dim fso As Object
    Dim ts As Object
    Dim strSigTextAll As String
    Dim strFontSize As String
    Dim intLocInFile As Integer
    Dim intLocOfEnd As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(PathToSig).openastextstream(1, -2)
    strSigTextAll = ts.readall
    
    If InStr(InStr(1, strSigTextAll, "<body"), strSigTextAll, ";font-family:", vbBinaryCompare) <> 0 Then
        intLocInFile = InStr(InStr(1, strSigTextAll, "<body"), strSigTextAll, "font-size:", vbBinaryCompare) + Len("font-size:")
        intLocOfEnd = InStr(intLocInFile, strSigTextAll, ".", vbBinaryCompare)
        strFontSize = Mid(strSigTextAll, intLocInFile, intLocOfEnd - intLocInFile)
    End If
    ts.Close
    GetSignatureFontSize = CInt(strFontSize)
    Set ts = Nothing
    Set fso = Nothing
End Function

Function ExtractTextInsideSpan(strHtmlText As String)

Dim boolStillFindingSpans As Boolean
Dim intLastSpanPos As Integer
Dim strNewSigText As String

boolStillFindingSpans = True
strNewSigText = "<p></p>"
intLastSpanPos = 0

While boolStillFindingSpans
    If InStr(intLastSpanPos + 1, strHtmlText, "<span", vbBinaryCompare) <> 0 Then
        intLastSpanPos = InStr(intLastSpanPos + 1, strHtmlText, "<span", vbBinaryCompare)
        intLastSpanPos = InStr(intLastSpanPos, strHtmlText, ">", vbBinaryCompare) + 1
        strNewSigText = strNewSigText & Mid(strHtmlText, intLastSpanPos, InStr(intLastSpanPos + 1, strHtmlText, "<", vbBinaryCompare) - intLastSpanPos) & "<p></p>"
    Else
        boolStillFindingSpans = False
    End If
Wend

Dim objCopyText As DataObject
Dim objCopyTextB As DataObject
Set objCopyText = New DataObject
    objCopyText.SetText strNewSigText
    objCopyText.PutInClipboard


On Error Resume Next
strNewSigText = Replace(strNewSigText, "@co.collin.tx.us<p", "@co.collin.tx.us</a><p")
On Error GoTo 0

Set objCopyTextB = New DataObject
    objCopyTextB.SetText strNewSigText
    objCopyTextB.PutInClipboard


ExtractTextInsideSpan = strNewSigText
End Function

Public Function UpdateWorkingForm()


Dim i, j As Integer
Dim PauseTime, PauseTimeText, Start, intFrmOpenTimer, StartTwo, Finish, TotalTime, intSecElapsed, intLimitInSec
Dim arrProgressIcons() As String
Dim boolEscape As Boolean
Dim v 'variant

On Error GoTo ErrHandleCodeHere
DoEvents
i = Me.btnQuit.BackColor
intFrmOpenTimer = Timer

ReDim arrProgressIcons(0 To 3)
arrProgressIcons(1) = "|"
arrProgressIcons(2) = "/"
arrProgressIcons(3) = "--"
arrProgressIcons(0) = "\"

j = -1
i = 1
Start = Timer ' Set start time.
intSecElapsed = 0
intLimitInSec = 5

boolEscape = False

With Me
    .Height = 60
    .btnQuit.Visible = False
    .lblSad.Visible = False
    .lblSorryOne.Visible = False
    .lblSorryTwo.Visible = False
    .BackColor = 16777215
End With

While boolEscape = False
    PauseTimeText = 0.5 ' longer pause for updating text progres spinner
    PauseTime = 0.01 ' Set duration.
    Me.Repaint
    DoEvents
    'If Timer < Start + PauseTimeText Then
        'DoEvents ' Yield to other processes.
    Me.lblProgressText.Caption = arrProgressIcons(i Mod 4)
    i = i + 1
    'End If
    Start = Timer
    
    Do While Timer < Start + PauseTimeText 'prevents text updating fast
        DoEvents
            strWindowSearchTitle = "- Message"
            boolFoundWindow = False
            Call SearchForWindowByTitle
            If boolFoundWindow = True Then
                Me.Hide
                'Unload frmWorking
                GoTo quitform
                boolEscape = True
            End If
        If Me.lblMovingBar.Left >= 12.5 And (Me.lblMovingBar.Left + Me.lblMovingBar.Width) <= 202.5 Then
            Me.lblMovingBar.Left = Me.lblMovingBar.Left + (2.5 * j)
        Else
            j = j * (-1)
            Me.lblMovingBar.Left = Me.lblMovingBar.Left + (2.5 * j)
        End If
        'Me.Repaint
        StartTwo = Timer
        Do While Timer < StartTwo + PauseTime 'gives a shorter pause time between moving progress bar updates
            DoEvents
        Loop
        intSecElapsed = Timer - intFrmOpenTimer
        If (Timer - intFrmOpenTimer) > 0.25 Then Me.lblSeconds.Caption = "Working now for " & CInt(intSecElapsed) & " seconds..."
        'Me.Repaint
    Loop
    
    If intSecElapsed >= intLimitInSec And intSecElapsed < (intLimitInSec + 4) Then
        With Me
            .Height = 175
            .btnQuit.Visible = True
            .lblSad.Visible = True
            .lblSorryOne.Visible = True
            .lblSorryTwo.Visible = True
            If intSecElapsed > intLimitInSec And intSecElapsed < (intLimitInSec + 3) Then
                If intSecElapsed Mod 2 = 0 Then
                    Me.BackColor = 192
                    For Each v In Me.Controls
                        With v
                            If .BackColor <> 65280 Then 'green
                                .BackColor = 192 'red
                                .ForeColor = 16777215 'white
                            End If
                        End With
                    Next v
                Else
                    Me.BackColor = 16777215
                    For Each v In Me.Controls
                        With v
                            If .BackColor <> 65280 Then
                                .BackColor = 16777215 'white
                                .ForeColor = -2147483630 'black
                            End If
                        End With
                    Next v
                End If
            Else
                Me.BackColor = 12632319 'light red
                For Each v In Me.Controls
                    With v
                        If .BackColor <> 65280 Then
                            .BackColor = 12632319 'white
                            .ForeColor = -2147483630 'black
                        End If
                    End With
                Next v
            End If
        End With
    End If


Wend

'Unload Me

Err.Number = 0
ErrHandleCodeHere:
If Err.Number = -2147418105 Then
    boolEscape = True
    Unload Me
End If

quitform:

End Function
