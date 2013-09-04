Attribute VB_Name = "m_HTML_Parse"
Option Explicit

Public Sub ParseTimesheetHTML()

    '   This sub is more for supervisors, who can't use the copy-paste method of importing from PeopleSoft

    Dim oFSO As New FileSystemObject
    Dim oFS
    Dim sText$, x$
    Dim boolInPunchBlock As Boolean

    boolInPunchBlock = False

    Set oFS = oFSO.OpenTextFile("C:\Users\e008922\Downloads\Report Time.htm")

    Do Until oFS.AtEndOfStream
        sText = oFS.ReadLine
        If StrComp(sText, "<tr valign='center'>", vbTextCompare) = 0 Then
            boolInPunchBlock = True
        End If
        If boolInPunchBlock And StrComp(sText, "</table>", vbTextCompare) = 0 Then
            boolInPunchBlock = False
        End If
        If boolInPunchBlock Then x = x & sText
    Loop

    Debug.Print x


End Sub
