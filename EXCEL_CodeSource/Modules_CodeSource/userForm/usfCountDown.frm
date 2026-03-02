VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCountDown 
   Caption         =   "Count down"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   OleObjectBlob   =   "usfCountDown.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCountDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'    Dim a%
'    Dim i%
'    Dim tempTime!
'
'    a = 10
'    lblCountDown.Caption = a
'    For i = 0 To 10
'    DoEvents
'
'    tempTime = Timer + 1
'    Do While tempTime > Timer
'    lblCountDown.Caption = a
'
'    Loop
'    a = a - 1
'
'    Next
Dim temps As Date
temps = Now()
Dim compte As Integer
compte = 10 'mettre le nombre de secondes ou en minutes
temps = DateAdd("s", compte, temps) 'pour des minutes, mettre "n" à la place de "s"
Do Until temps < Now()
DoEvents
    usfCountDown.lblCountDown.Caption = Format((temps - Now()), "hh:mm:ss")
'ActivePresentation.Slides(1).Shapes("affichage").TextFrame.TextRange = Format((temps - Now()), "hh:mm:ss")
Loop
End Sub


