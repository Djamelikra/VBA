Attribute VB_Name = "mSound"
Option Explicit

'Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByValdwFlags As Long) As Boolean

Private Declare PtrSafe Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName$, _
        ByVal hModule&, ByVal dwFlags&)
 
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const NULL_POINTER = 0&

 Sub soundStarter()
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
 Sub soundStarterWmp()
'To play a sound when starting the application
    Dim fileSoundName$ 'sound file name
    Dim dirSoundName$ 'sound directory name
    Dim fullFilePath$ 'sound file full name
    Dim sep$
 
    sep = Application.PathSeparator
    
    dirSoundName = "Sound"
    fileSoundName = "missionImpossible.mp3"
    fullFilePath = ThisWorkbook.Path & sep & dirSoundName & sep & fileSoundName
'        fullFilePath = fetchCurrentDirectory & sep & dirSoundName & sep & fileSoundName

    
'    Call openMediaPlayer
    Call PlaySound(fullFilePath, NULL_POINTER, SND_ASYNC Or SND_FILENAME)
    
End Sub


 Sub soundEnd()
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

Sub JouerSon()
Dim MonWav As String
Dim myWav As String
    MonWav = "C:\LeSon.wav"     '... chemin et nom à adapter
    myWav = CurDir$
'    call PlaySound
    Call PlaySound(MonWav, 0&, SND_ASYNC Or SND_FILENAME)
End Sub

Sub testFullName()
'    MsgBox ThisWorkbook.FullName, vbInformation, "FullName"
    Debug.Print ThisWorkbook.FullName
    
'        MsgBox ThisWorkbook.Path, vbInformation, "Path"
     Debug.Print ThisWorkbook.Path
    
'    MsgBox ThisWorkbook.Name, vbInformation, "Name"
     Debug.Print ThisWorkbook.Name
     
     With Selection
        .TypeText ThisWorkbook.FullName & vbNewLine
        .TypeText ThisWorkbook.Path & vbNewLine
        .TypeText ThisWorkbook.Name & vbNewLine
     End With

    
End Sub
Sub testPath()
    MsgBox ThisWorkbook.Path
    
End Sub
Sub testCurDir()
    MsgBox CurDir$
    
End Sub
Function getCurrentDirectory() As String
    Dim sPath$
    Dim sFolder As Variant
    
    If ThisWorkbook.Path = vbNullString Then
        MsgBox "You must save your document, before !."
        Exit Function
    End If
    sPath = ThisWorkbook.Path
    sFolder = Split(sPath, Application.PathSeparator)
    sFolder = sFolder(UBound(sFolder))
    getCurrentDirectory = sFolder

End Function
Sub testGetDir()
    MsgBox getCurrentDirectory
End Sub
Sub nom_dossier()
Dim chemin As String
Dim dossier As Variant

If ThisWorkbook.Path = vbNullString Then
MsgBox "Commencez par enregistrer votre document."
Exit Sub
End If
chemin = ThisWorkbook.Path
dossier = Split(chemin, Application.PathSeparator)
dossier = dossier(UBound(dossier))
MsgBox dossier

Selection.TypeText dossier

End Sub
