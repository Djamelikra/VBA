Attribute VB_Name = "mcolorIndex"
Option Explicit

Sub RefCouleur()

Dim x As Integer

For x = 1 To 56
  If x < 29 Then
    Cells(x, 1).Interior.ColorIndex = x
    Cells(x, 2) = x
  Else
    Cells(x - 28, 3).Interior.ColorIndex = x
    Cells(x - 28, 4) = x
  End If
Next x

End Sub

