VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfSplashScreen 
   Caption         =   "Applications"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11070
   OleObjectBlob   =   "usfSplashScreenAppOriginal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'hide the close button of the UserForm:
Private Const SC_CLOSE = &HF060& '61536
Private Const MF_BYCOMMAND = &H0& '0

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
#Else
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
#End If
  
#If VBA7 Then
Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Sub UserForm_Initialize()
Dim hSysMenu As Long
Dim MeHwnd As Long
    MeHwnd = FindWindowA(vbNullString, Me.Caption)
    If MeHwnd > 0 Then
        hSysMenu = GetSystemMenu(MeHwnd, False)
        RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
    Else
        MsgBox "Handle de " & Me.Caption & " Introuvable", vbCritical
    End If
End Sub
Private Sub UserForm_Activate()
'''Dim PauseTime&
'''Dim Start&
'''Dim Finish&
'''Dim TotalTime&


'''Call playMi

'###########################################################
'''    PauseTime = 51   ' Définit la durée.
'''    Start = Timer    ' Définit l'heure de début.
'''    Do While Timer < Start + PauseTime
'''        DoEvents    ' Donne le contrôle à d'autres processus.
'''    Loop
'''    Finish = Timer    ' Définit l'heure de fin.
'''
'''        Unload usfSplashScreen
 
 '################################################
 'The countdown algorithm
    Dim a As Long
    Dim i As Long
    Dim timer1 As Single

    Call playMi
    
    a = 49
    lblCountDown.Caption = a
    'Modif
    For i = -1 To a '-> Modif
        DoEvents
        timer1 = Timer + 1
        Do While timer1 > Timer
            With Me.lblCountDown
                .Caption = a
                .AutoSize = False
            End With
            If Me.lblCountDown < 2 Then
                With Me.lblSecond
                    .Caption = "second"
                    .AutoSize = True
                End With
            End If
        Loop
        a = a - 1
    Next
    
      ThisWorkbook.Application.Visible = True
    
    Unload Me
 
End Sub

