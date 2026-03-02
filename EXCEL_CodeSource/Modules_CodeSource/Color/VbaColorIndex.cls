VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaColorIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub vbaColorIndexRef()

Dim rng As Range, cell As Range
Dim i As Integer
 
i = 1
 
'56 cells to hold all default palette colors
Set rng = Range("A1:G8")
 
For Each cell In rng
 
    cell.Interior.ColorIndex = i
 
    cell.Value = i
 
    i = i + 1
 
Next cell

End Sub
