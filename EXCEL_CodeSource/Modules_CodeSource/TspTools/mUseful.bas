Attribute VB_Name = "mUseful"
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
    
    Beep
    
End Sub
Sub msgBoxTempMsgWait(sMessage As String, secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup(sMessage, secondWait, gAppName, vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
    
    Beep
    
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
Sub currentFileSize()
   Dim strMonFichier$
   Dim lngSizeFich&
'   strMonFichier = Application.GetOpenFilename("Fichiers texte (*.txt), *.txt, Fichiers Excel (*.xls), *.xls", , "Sélection des fichiers à ouvrir: ")

    strMonFichier = ThisWorkbook.FullName

   lngSizeFich = Len(strMonFichier)
   MsgBox "Size of the file: " & strMonFichier & " : " & _
                  vbCr & lngSizeFich & " Octets", vbInformation, "File size :"
End Sub
    
    ''***********************FileLen***********************************************
'**Renvoie une valeur de type Long indiquant la taille en octets d'un fichier.
'**FileLen(pathname$) As long

Sub TailleFichier()
   Dim strMonFichier$
   Dim lngSizeFich&
   strMonFichier = Application.GetOpenFilename("Fichiers texte (*.txt), *.txt, Fichiers Excel (*.xls), *.xls", , "Sélection des fichiers à ouvrir: ")

   lngSizeFich = Len(strMonFichier)
   MsgBox "Size of the file: " & strMonFichier & " : " & _
                  vbCr & lngSizeFich & " Octets", vbInformation, "File size :"
End Sub
'******'**********************LOF ****************************************************
'***************************(LOF = len of file)********************************************
'**Renvoie une valeur de type Long représentant la taille, exprimée en octets, d'un fichier ouvert à l'aide de l'instruction Open.
Sub TailleFichOuvert()
   Dim SizeFichier&
   Dim lngNombreCaract&
   Open "D:\DEVELOPPEMENT\PROJECTS\002_ALGORITHMIQUE\MONFICH.txt" For Input As #1
   ' Ouvre le fichier.
   SizeFichier = LOF(1)    ' Lit la taille du fichier.
   MsgBox SizeFichier & " Octets", vbInformation, "File size"
   Close #1    ' Ferme le fichier.
End Sub

Sub openWorkspace()
    Dim workPath As String
    workPath = ThisWorkbook.Path
    MsgBox workPath, vbInformation, "The Workspace Directory: "
    Shell "C:\windows\explorer.exe " & workPath, vbMaximizedFocus
End Sub
