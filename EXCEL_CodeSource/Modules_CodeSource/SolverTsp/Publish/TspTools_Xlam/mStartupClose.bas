Attribute VB_Name = "mstartupClose"
Option Explicit

Sub Auto_open()
    usfSheets.Show
'    usfImportExportCsv.Show
  ActiveWindow.WindowState = xlMaximized
   Beep
   
    
   


  Call openWorkspace


End Sub

Sub Auto_Close()

    MsgBox "Eurék@!", vbInformation, gAppName & " " & gCR
      
    Call applauseSound
End Sub
Sub openWorkspace()
    Dim workPath As String
    workPath = ThisWorkbook.Path
    MsgBox workPath, vbInformation, "The Workspace Directory: "
    Shell "C:\windows\explorer.exe " & workPath, vbMaximizedFocus
    Beep
End Sub
