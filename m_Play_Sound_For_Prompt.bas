Attribute VB_Name = "m_Play_Sound_For_Prompt"
Option Explicit

Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const ENUM_CURRENT_SETTINGS As Long = -1
Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32

Private Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * CCHDEVICENAME
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


Private Declare Function EnumDisplayDevices Lib "user32.dll" Alias "EnumDisplayDevicesA" (ByVal lpDevice As String, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Long











Sub mPlaySoundForPrompt(ByVal WhatSound As String, Optional Flags As Long = 0)
    If Dir(WhatSound, vbNormal) = "" Then
        ' WhatSound is not a file. Get the file named by
        ' WhatSound from the Windows\Media directory.
        WhatSound = Environ("SystemRoot") & "\Media\" & WhatSound
        If InStr(1, WhatSound, ".") = 0 Then
            ' if WhatSound does not have a .wav extension,
            ' add one.
            WhatSound = WhatSound & ".wav"
        End If
        If Dir(WhatSound, vbNormal) = vbNullString Then
            ' Can't find the file. Do a simple Beep.
            Beep
            Exit Sub
        End If
    Else
        ' WhatSound is a file. Use it.
    End If
    ' Finally, play the sound.
    sndPlaySound32 WhatSound, Flags
End Sub


Public Sub soundTest()



    Dim indAdapter As Long, indDisplay As Long
    Dim ddAdapters As DISPLAY_DEVICE, ddDisplays As DISPLAY_DEVICE
    ddAdapters.cb = Len(ddAdapters): ddDisplays.cb = Len(ddDisplays)

    indAdapter = 0
    Do Until EnumDisplayDevices(vbNullString, indAdapter, ddAdapters, 0) = 0

        If (ddAdapters.StateFlags And DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) = DISPLAY_DEVICE_ATTACHED_TO_DESKTOP Then

            Dim NullCharPos As Long
            NullCharPos = InStr(ddAdapters.DeviceName, vbNullChar)

            Dim CurDeviceName As String

            If NullCharPos > 0 Then
                CurDeviceName = Left$(ddAdapters.DeviceName, NullCharPos - 1)
            Else
                CurDeviceName = ddAdapters.DeviceName
            End If

            Dim dmode As DEVMODE
            dmode.dmSize = Len(dmode)

            EnumDisplaySettings CurDeviceName, ENUM_CURRENT_SETTINGS, dmode

            MsgBox "Width: " & dmode.dmPelsWidth
            MsgBox "Height: " & dmode.dmPelsHeight

        End If

        indAdapter = indAdapter + 1
    Loop


    Call mPlaySoundForPrompt("Windows Hardware Fail.wav", &H1)


End Sub
