Attribute VB_Name = "mApi_mciSendString"
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias _
   "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
   lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
   hwndCallback As Long) As Long

Private musicFile$
Dim Play As Variant

Public Sub playAnySound(ByVal File$)
On Error GoTo errHandler

musicFile = File    'path has been included

If Play <> vbNull Then 'this triggers if can't play the file
   Play = mciSendString("play " & musicFile, 0&, 0, 0) 'i tried this aproach and it works
End If

Exit Sub
errHandler:
     MsgBox "The following error has occurred :" & vbCrLf _
                        & "Error number:  " & Err.Number & vbCrLf _
                        & "Type of error :  " & Err.Description, vbCritical

End Sub
Public Sub playMi()
On Error GoTo errHandler
'To play a sound when starting the application
    Dim fileSoundName$ 'sound file name
    Dim fullFilePath$ 'sound file full name
    Dim sep$
    sep = Application.PathSeparator

   If ThisWorkbook.Path = "" Then
        MsgBox "Is your file has been saved ?", vbQuestion
        Exit Sub
    End If
    
    

    fileSoundName = "missionImpossible.mp3"
    fullFilePath = sep & fileSoundName
    
    ChDir (ThisWorkbook.Path)
    ChDir "Sound"
'    MsgBox CurDir
    
    fullFilePath = CurDir & sep & fileSoundName
    
'      MsgBox fullFilePath
      Call playAnySound(fullFilePath)
      
      Exit Sub
errHandler:
     MsgBox "The following error has occurred :" & vbCrLf _
                        & "Error number:  " & Err.Number & vbCrLf _
                        & "Type of error :  " & Err.Description, vbCritical

    
End Sub

Sub test_playAnySound()
    playAnySound ("D:\missionImpossible.mp3")
End Sub
Public Sub StopSound(Optional ByVal FullFile$)
Play = mciSendString("close " & musicFile, 0&, 0, 0)
End Sub
Sub testStopSound()
    StopSound ("D:\missionImpossible.mp3")
End Sub


