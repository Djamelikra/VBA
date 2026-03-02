Attribute VB_Name = "mWMI"
Option Explicit
'mWMI
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassname As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
 
' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
'to do

Sub ActiveProcess()
Dim svc As Object
Dim sQuery As String
Dim oproc
Dim Idwind
 
Set svc = GetObject("winmgmts:root\cimv2")
sQuery = "select * from win32_process"
 
For Each oproc In svc.execQuery(sQuery)
    Debug.Print oproc.Name
    If oproc.Name = "notepad.exe" Then
        Idwind = "Sans titre"
    End If
Next
    If Idwind = "" Then
        Idwind = Shell("C:\Windows\System32\notepad.exe", vbNormalFocus)
        AppActivate Idwind
    Else
        Dim toto As String
        Dim hwnd As Long
        toto = "Untitled - Notepad"
        hwnd = FindWindow(vbNullString, toto)
        ' Hwnd = FindWindow("OpusApp", vbNullString)
 
        If hwnd = 0 Then Exit Sub
        SetForegroundWindow hwnd
        ShowWindow hwnd, SW_SHOWMAXIMIZED
 
    End If
Set svc = Nothing
End Sub
