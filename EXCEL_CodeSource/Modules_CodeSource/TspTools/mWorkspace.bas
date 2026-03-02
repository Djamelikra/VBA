Attribute VB_Name = "mWorkspace"
Option Explicit
'mWorkspace

Sub openFolder()
    Dim Chemin As String
 
    Chemin = "C:\Temp\"
    Shell "C:\windows\explorer.exe " & Chemin, vbMaximizedFocus
End Sub
Function openTheDirectory(dirPath As String) As Boolean
    'dirPath = directory path
'     Dim oShell As Shell
'    Set oShell = New Shell
'    oShell.Explore (dirPath)
End Function
Sub openWorkspace()
    Dim workPath As String
    workPath = ThisWorkbook.Path
    MsgBox workPath, vbInformation, "The Workspace Directory: "
    Shell "C:\windows\explorer.exe " & workPath, vbMaximizedFocus
End Sub
