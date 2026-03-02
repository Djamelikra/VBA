Attribute VB_Name = "mOpenFolder"
Option Explicit

Sub repertoireExplorateur_V3()
    Dim strFolder As String
    strFolder = ThisWorkbook.Path
 
'  ThisWorkbook.FollowHyperlink strFolder
  
  
  'Shell
  Shell "C:\windows\explorer.exe " & strFolder, vbNormalFocus
  
End Sub

Sub repertoireExplorateur_V2()
    Dim Chemin As String
 
    Chemin = "C:\Documents and Settings\dossier"
    ThisWorkbook.FollowHyperlink Chemin
End Sub
Sub OpenFolder()


    'Nécessite d'activer la référence "Microsoft Shell Controls and Automation"
    Dim objShell As Shell
     Dim strFolder As String
 
    strFolder = ThisWorkbook.Path
 
    Set objShell = New Shell
    objShell.Explore (strFolder)


End Sub
Sub OpenMyDirectory2()
Dim MyFolder As String
MyFolder = ThisWorkbook.Path

'vérifie si le Dossier existe
If Len(Dir(MyFolder, vbDirectory)) > 0 Then
   Shell Environ("WINDIR") & "\explorer.exe " & MyFolder, vbNormalFocus
   Else
   MsgBox "OUPS !", vbExclamation, "There isn't such folder !"
End If

End Sub
Public Function OpenDirectory(MyFolder As String)

'vérifie si le Dossier existe
If Len(Dir(MyFolder, vbDirectory)) > 0 Then
     Shell Environ("WINDIR") & "\explorer.exe " & MyFolder, vbNormalFocus
      Else
   MsgBox "OUPS !", vbExclamation, "There isn't such folder !"
End If

End Function
Sub test()

    Dim strFolder As String
'    strFolder = ThisWorkbook.FullName
    strFolder = ThisWorkbook.Path
    
    
    MsgBox strFolder
    
    

End Sub
