Attribute VB_Name = "mSound2"
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias _
   "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
   lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
   hwndCallback As Long) As Long

Private sMusicFile As String
Dim Play

Public Sub Sound2(ByVal File$)

sMusicFile = File    'path has been included. Ex. "C:\3rdMan.mp3

Play = mciSendString("play " & sMusicFile, 0&, 0, 0)
If Play <> vbNull Then 'this triggers if can't play the file
    Play = mciSendString("play " & sMusicFile, 0&, 0, 0) 'i tried this aproach, and it works
End If
   
End Sub


Public Sub StopSound(Optional ByVal FullFile$)
Play = mciSendString("close " & sMusicFile, 0&, 0, 0)
End Sub
