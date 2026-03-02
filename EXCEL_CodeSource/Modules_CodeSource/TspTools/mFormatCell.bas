Attribute VB_Name = "mFormatCell"
Option Explicit

Sub CellBackgdYellow()
Attribute CellBackgdYellow.VB_Description = "Cell's background yellow"
Attribute CellBackgdYellow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgrdGreenSolutionTsp()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(51, 255, 143)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub selectAutofit()
'autofit selection data
    Selection.EntireColumn.AutoFit
    ActiveCell.Select
End Sub
