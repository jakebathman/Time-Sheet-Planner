Attribute VB_Name = "Window_Search"
Option Explicit

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
   lpRect As RECT) As Long

Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
   ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function EnumWindows Lib "user32" _
   (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent _
   As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId _
   As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Declare Function GetWindowThreadProcessId Lib "user32" _
   (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
   (ByVal hwnd As Long, ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
   (ByVal hwnd As Long, ByVal lpString As String, _
   ByVal cch As Long) As Long

Public TopCount As Integer     ' Number of Top level Windows
Public ChildCount As Integer   ' Number of Child Windows
Public ThreadCount As Integer  ' Number of Thread Windows


' code from: http://support.microsoft.com/kb/183009
Public Function SearchForWindowByTitle()
    Dim lRet As Long, lParam As Long
    Dim lhWnd As Long
    
    'lhWnd = Me.hwnd  ' Find the Form's Child Windows
    ' Comment the line above and uncomment the line below to
    ' enumerate Windows for the DeskTop rather than for the Form
    lhWnd = GetDesktopWindow()  ' Find the Desktop's Child Windows
    ' enumerate the list
    lRet = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
    
    If lRet = 0 Then boolFoundWindow = True Else boolFoundWindow = False
    
    
'    Dim lRet As Long
'    Dim lParam As Long
    
    'enumerate the list
    'lRet = EnumWindows(AddressOf EnumWinProc, lParam)
    ' How many Windows did we find?
'    Debug.Print TopCount; " Total Top level Windows"
'    Debug.Print ChildCount; " Total Child Windows"
'    Debug.Print ThreadCount; " Total Thread Windows"
'    Debug.Print "For a grand total of "; TopCount + ChildCount + ThreadCount; " Windows!"
End Function
      


Function EnumWinProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long
   Dim RetVal As Long, ProcessID As Long, ThreadID As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String
    DoEvents
   RetVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)
   TopCount = TopCount + 1
   ' see the Windows Class and Title for each top level Window
   '''Debug.Print "Top level Class = "; WinClass; ", Title = "; WinTitle
   ' Usually either enumerate Child or Thread Windows, not both.
   ' In this example, EnumThreadWindows may produce a very long list!
   RetVal = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
   ThreadID = GetWindowThreadProcessId(lhWnd, ProcessID)
   RetVal = EnumThreadWindows(ThreadID, AddressOf EnumThreadProc, _
   lParam)
   EnumWinProc = True
End Function

Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
   Dim RetVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String
   Dim WinRect As RECT
   Dim WinWidth As Long, WinHeight As Long
    DoEvents
   RetVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)
   ChildCount = ChildCount + 1
   ' see the Windows Class and Title for each Child Window enumerated
   '''Debug.Print "   Child Class = "; WinClass; ", Title = "; WinTitle
   ' You can find any type of Window by searching for its WinClass
   If InStr(1, WinTitle, strWindowSearchTitle, vbTextCompare) Then ' TextBox Window
        'MsgBox ("Found window that contains ""Timesheet""!")
        DoEvents
        EnumChildProc = False
   Else
        EnumChildProc = True
   End If
End Function

Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
   Dim RetVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String
    DoEvents
   RetVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)
   ThreadCount = ThreadCount + 1
   ' see the Windows Class and Title for top level Window
   '''Debug.Print "Thread Window Class = "; WinClass; ", Title = ";  WinTitle
   EnumThreadProc = True
End Function

Public Function StripNulls(OriginalStr As String) As String
   ' This removes the extra Nulls so String comparisons will work
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function




