Attribute VB_Name = "mPowershell"
Option Explicit
'mPowershell
Private Function openPowerShell() As Boolean
    Dim sClassName As String
    Dim sWindowName As String
    Dim lHwnd As Long
    Dim bExe As Boolean
    Dim sDirPs As String
    Dim lResult As Long
    
    'init
    openPowerShell = False
    
    'retrieve the window calc
    sClassName = vbNullString
    sWindowName = ("mspaint")
    lHwnd = FindTheWindow(sClassName, sWindowName)
    
    'if not found so retrieve calc directory
    If lHwnd = 0 Then
        sDirPs = String$(255, 0)
        lResult = FindTheExecutable("powershell.exe", "C:\Windows\System32\WindowsPowerShell\v1.0\", sDirPs)
            If lResult = 0 Then
                MsgBox "Sorry ! Paint app not found !", vbExclamation
                Exit Function
            Else: Shell sDirPs, vbNormalFocus
                   BringWindowToTop lHwnd
                    'Affiche en mode "Normal"
'                    ShowWindow lHwnd, 1
'                MsgBox "The Paint application has been launched.", vbInformation
            End If
    Else
         MsgBox "The Powershell application is already active.", vbInformation
          
            
    End If
    
    openPowerShell = True
End Function

