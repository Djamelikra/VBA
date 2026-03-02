Attribute VB_Name = "mUsual"
Option Explicit

Private Declare PtrSafe Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName$, _
        ByVal hModule&, ByVal dwFlags&)
 
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const NULL_POINTER = 0&
Sub msgBoxTemp(secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup("Click OK (this window closes automatically after)." & " " & secondWait & " " & "seconds.", _
    secondWait, gAppName, vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Sub CellBackgdYellow()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
 Sub applauseSound()
'To play a sound when starting the application
    Dim fileSoundName$ 'sound file name
    Dim dirSoundName$ 'sound directory name
    Dim fullFilePath$ 'sound file full name
    Dim sep$
    
    sep = Application.PathSeparator
    
    dirSoundName = "Sound"
    fileSoundName = "applaud.wav"
    
    fullFilePath = ThisWorkbook.Path & sep & dirSoundName & sep & fileSoundName
    
    Call PlaySound(fullFilePath, NULL_POINTER, SND_ASYNC Or SND_FILENAME)
    
End Sub
 Sub tadaSound()
'To play a sound when starting the application
    Dim fileSoundName$ 'sound file name
    Dim dirSoundName$ 'sound directory name
    Dim fullFilePath$ 'sound file full name
    Dim sep$
    
    sep = Application.PathSeparator
    
    dirSoundName = "C:\WINDOWS\MEDIA"
    fileSoundName = "TADA.wav"
    

    fullFilePath = dirSoundName & sep & fileSoundName
    
    Call PlaySound(fullFilePath, NULL_POINTER, SND_ASYNC Or SND_FILENAME)
    End Sub
    
