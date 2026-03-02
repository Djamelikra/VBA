VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   OleObjectBlob   =   "usfMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim X As Single
Dim Y As Single
 
 
'Création de la barre d'outils lors du lancement du UserForm
Private Sub UserForm_Initialize()
    Dim Barre As CommandBar
 
    Set Barre = CommandBars.Add("MenuUSF", msoBarPopup, False, True)
 
    With Barre.Controls.Add(msoControlButton, 1, , , True)
        .Caption = "Menu 01"
        .FaceId = 50
        'La procédure va appeler une macro nommée "Macro1", lorsque vous cliquerez
        'sur le bouton.
        .OnAction = "Macro1"
    End With
 
    With Barre.Controls.Add(msoControlButton, 2, , , True)
        .Caption = "Menu 02"
        .FaceId = 49
        'La procédure va appeler une macro nommée "Macro2", lorsque vous cliquerez
        'sur le bouton.
        .OnAction = "Macro2"
    End With
 
 
    With Me
        X = (.Width - .InsideWidth) / 2 + 8
        Y = .Height - .InsideHeight - X + 24
    End With
End Sub
 
 
 
'Affiche la barre d'outils lorsque vous cliquez sur le label.
Private Sub Label1_Click()
    Dim PosX As Single, PosY As Single
 
    PosX = (Me.Left + X + Label1.Left) * 4 / 3
    PosY = (Me.Top + Y + Label1.Top) * 4 / 3
 
    Application.CommandBars("MenuUSF").ShowPopup PosX, PosY
End Sub
 
 
 
'Supprime la barre d'outils lors de la fermeture du UserForm
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    CommandBars("MenuUSF").Delete
End Sub
