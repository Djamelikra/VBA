VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfRefEdit 
   Caption         =   "Select range"
   ClientHeight    =   720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "usfRefEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfRefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   Dim Plage As String
    
    Plage = RefEdit1.Value
    'Vérifie s'il y a eu une sélection
    If Plage = "" Then
        MsgBox "Opération annulée"
        Exit Sub
    End If
    
    'Insère une croix dans toutes les cellules de la plage
    Range(Plage) = "x"
    Unload Me
End Sub

