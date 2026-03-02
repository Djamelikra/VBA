Attribute VB_Name = "mAppTop"
Option Explicit
 
Public Declare PtrSafe Function BringWindowToTop Lib "user32" _
    (ByVal Hwnd As Long) As Long
 
Public Declare PtrSafe Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
 
Public Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
 
 
'Remarque importante:
'La procédure ne doit pas être déclenchée depuis l'éditeur de macros /!\
'
Sub ApplicationPremierPlan()
    Dim Hwnd As Long
 
    'Récupère le Handle d'une fenêtre (la calculatrice dans cet exemple).
    'Le Handle est un nombre entier unique généré par Windows afin d'identifier les fenêtres.
    '"Calculatrice" correspond au titre de la fenêtre.
    Hwnd = FindWindow(vbNullString, "Calculatrice")
 
    'Si la calculatrice est déjà ouverte
    If Hwnd > 0 Then
        'Ramène la calculatrice au premier plan
        BringWindowToTop Hwnd
        'Affiche en mode "Normal"
        ShowWindow Hwnd, 1
        Else
        'Sinon, ouvre la calculatrice
        Shell "C:\WINDOWS\system32\calc.exe", vbNormalFocus
    End If
End Sub
