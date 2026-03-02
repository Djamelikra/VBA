Attribute VB_Name = "mstartupClose"
Option Explicit
Private Const ALGO_DOC As String = "VerbatimAlgoV3.docx"
Private Const ALGO_XL As String = "Algorithm.xlsb"

Sub Auto_open()
'    usfSheets.Show
    '    usfImportExportCsv.Show
    ActiveWindow.WindowState = xlMaximized
    Beep
    Call openWorkspace
    Call OpenVerbatim
    Beep
    Call OpenAlgorithm
    Beep
    

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


Sub OpenVerbatim()
'    Dim strFileName As String
    Dim strFilePath As String
    
'    strFileName = "VerbatimAlgo.docx"
    strFilePath = ThisWorkbook.Path & Application.PathSeparator & ALGO_DOC
    ThisWorkbook.FollowHyperlink (strFilePath)
End Sub
Sub OpenAlgorithm()
'    Dim strFileName As String
    Dim strFilePath As String
    
'    strFileName = "Algorithm.xlsb"
    strFilePath = ThisWorkbook.Path & Application.PathSeparator & ALGO_XL
    ThisWorkbook.FollowHyperlink (strFilePath)
End Sub

Sub openUsfSheets()
    usfSheets.Show
End Sub
