VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfSplashScreen 
   Caption         =   "Applications"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "usfSplashScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Activate()
Dim PauseTime&
Dim Start&
Dim Finish&
Dim TotalTime&

Call playMi

    PauseTime = 51   ' Définit la durée.
    Start = Timer    ' Définit l'heure de début.
    Do While Timer < Start + PauseTime
        DoEvents    ' Donne le contrôle à d'autres processus.
    Loop
    Finish = Timer    ' Définit l'heure de fin.
    

'    usfSplashScreen.hide
        Unload usfSplashScreen
 
 
End Sub
