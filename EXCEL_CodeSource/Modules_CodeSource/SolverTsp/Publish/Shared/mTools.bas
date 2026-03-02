Attribute VB_Name = "mTools"
Option Explicit
'###########################################################################################
Sub Auto_open()
 
With Application
    .DisplayFullScreen = True
End With
End Sub

Sub Mandatory()

    With Selection.Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Sub


Sub DisplaySheetAlgo()
'
' DisplaySheetAlgo Macro
    Sheets("Algo").Select
    Range("C31").Select
End Sub



Sub DoubleStar()
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDashDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDashDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDashDot
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub





Sub Ordre0_HCP()
Attribute Ordre0_HCP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Ordre0_HCP Macro
'

'
    Range("B76:N88").Select
    Selection.Copy
    Range("B92").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A102").Select
End Sub
Sub IsSymetric()
'to do
'launch_euclidianNorm2D

End Sub
Sub PEP()
'
' PEP Macro
' Pauli Exclusion Principle
'

'
'    Range("M107:N107").Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

