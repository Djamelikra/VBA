Attribute VB_Name = "mStartUp"
Option Explicit

Sub Auto_open()

  Call openWorkspace


'With Application
'    .DisplayFullScreen = True
'End With

'     usfTSP.Show

End Sub
Sub openWorkspace()
    Dim workPath As String
    workPath = ThisWorkbook.Path
    MsgBox workPath, vbInformation, "The Workspace Directory: "
    Shell "C:\windows\explorer.exe " & workPath, vbMaximizedFocus
    Beep
End Sub
