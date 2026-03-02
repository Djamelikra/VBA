Attribute VB_Name = "mWindowsPowershell"
Option Explicit
'mWindowsPowershell
Public Declare PtrSafe Function FindTheWindow Lib "user32" Alias "FindWindowA" ( _
                 ByVal lpClassname As String, _
                 ByVal lpWindowName As String) As Long
                 

Public Declare PtrSafe Function FindTheExecutable Lib "shell32.dll" Alias "FindExecutableA" ( _
                 ByVal lpFile As String, _
                 ByVal lpDirectory As String, _
                 ByVal lpResult As String) As Long
                 
                 
Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
    
    
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'#Windows Powershell lambda
Private Function openWindowsPowershell() As Boolean
    Dim sClassName As String
    Dim sWindowName As String
    Dim lHwnd As Long
    Dim bExe As Boolean
    Dim sDirPs As String
    Dim lResult As Long
    
    'init
    openWindowsPowershell = False
    
    'retrieve the window calc
    sClassName = vbNullString
    sWindowName = ("Windows Powershell")
    lHwnd = FindTheWindow(sClassName, sWindowName)
    
    'if not found so retrieve Windows Powershell directory
    If lHwnd = 0 Then
        sDirPs = String$(255, 0)
        lResult = FindTheExecutable("powershell.exe", "C:\Windows\System32\WindowsPowerShell\v1.0\", sDirPs)
            If lResult = 0 Then
                MsgBox "Sorry ! Windows Powershell application not found !", vbExclamation
                Exit Function
            Else: Shell sDirPs, vbNormalFocus
'                   BringWindowToTop lHwnd
                    'Affiche en mode "Normal"
                    ShowWindow lHwnd, 1
'                MsgBox "The Windows Powershell application has been launched.", vbInformation
            End If
    Else
         MsgBox "The Windows Powershell application is already active.", vbInformation
          
            
    End If
    
    openWindowsPowershell = True
End Function

Sub launchWindowsPowershell()
    Call openWindowsPowershell
    
End Sub
'#Windows Powershell Administrateur
Private Function openWindowsPowershellAsAdmin() As Boolean
On Error GoTo HandleError

    Dim sClassName As String
    Dim sWindowName As String
    Dim lHwnd As Long
    Dim bExe As Boolean
    Dim sDirPs As String
    Dim lResult As Long
    
    'init
    openWindowsPowershellAsAdmin = False
    
'    Call openBatchForPowershell
'
'    Application.Wait (Now + TimeValue("00:00:01"))
'    'Variant
''    Application.Wait (Now + TimeSerial(0, 0, 1))
    
    
    'TO DO
    
    'retrieve the window calc
    sClassName = vbNullString
    sWindowName = ("Administrateur : Windows Powershell")
    lHwnd = FindTheWindow(sClassName, sWindowName)
    
                                            '    'if not found so retrieve Windows Powershell directory
                                            '    If lHwnd = 0 Then
                                            '        sDirPs = String$(255, 0)
                                            '        lResult = FindTheExecutable("powershell.exe", "C:\Windows\System32\WindowsPowerShell\v1.0\", sDirPs)
                                            '            If lResult = 0 Then
                                            '                MsgBox "Sorry ! Windows Powershell application not found !", vbExclamation
                                            '                Exit Function
                                            '            Else: Shell sDirPs, vbNormalFocus
                                            ''                   BringWindowToTop lHwnd
                                            '                    'Affiche en mode "Normal"
                                            '                    ShowWindow lHwnd, 1
                                            ''                MsgBox "The Windows Powershell application has been launched.", vbInformation
                                            '            End If
                                            '    Else
                                            '         MsgBox "The Windows Powershell application is already active.", vbInformation
                                            '    End If
                                            
     'if not found so retrieve Windows Powershell directory
    If lHwnd = 0 Then
        Call openBatchForPowershell
      Else
    MsgBox "The Windows Powershell application is already active.", vbInformation
        Exit Function
    Application.Wait (Now + TimeValue("00:00:01"))
    'Variant
'    Application.Wait (Now + TimeSerial(0, 0, 1))
    
    End If
                                            
    
    openWindowsPowershellAsAdmin = True
    
HandleExit:
    Exit Function
HandleError:
      MsgBox "Error n° " & Err.Number & " Occured" & vbNewLine & Err.Description & vbNewLine & Err.Source, vbExclamation, Err.Source
    Resume HandleExit
    
End Function

Private Sub openBatchForPowershell()
    'Mandatory
    Call Shell("OpenPsAsAdmin.bat")
    
End Sub

Sub launchWindowsPowershellAdmin()
    Call openWindowsPowershellAsAdmin
    
End Sub






