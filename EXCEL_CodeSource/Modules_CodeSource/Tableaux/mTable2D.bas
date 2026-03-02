Attribute VB_Name = "mTable2D"
Option Explicit
Sub multiplTable()
'^ De -9 223 372 036 854 775 808 à 9 223 372 036 854 775 807
'& -2 147 483 648 to 2 147 483 648
'% -32 768 to 32 767
'! -3.402823E38 to -1.401298E-45 for negative values, 1.401298E-45 to 3.402823E38 for positive values
'@ -922 337 203 685 477.5808 to 922 337 203 685 477.5807
'# -1.79769313486232e+308 to -4.94065645841247E-324 for negative values, 4.94065645841247E-324 to 1.79769313486232e+308 for positive values.
'2D table
    Dim iLine&
    Dim jCol&
    
    Cells.Clear
      
    For iLine = 1 To 10
        For jCol = 1 To 10
            Cells(iLine, jCol) = iLine * jCol
        Next
    Next
End Sub

Sub tableHeaders()
'2D table
    Dim iLine&
    Dim jCol&
    
    Cells.Clear
    
    For iLine = 1 To 10
        Cells(iLine + 1, 1) = iLine
        Cells(iLine + 1, 1).Font.Bold = True
        Cells(1, iLine + 1) = iLine
        Cells(1, iLine + 1).Font.Bold = True
        For jCol = 1 To 10
            Cells(iLine + 1, jCol + 1) = iLine * jCol
        Next
    Next
End Sub

Sub TableMulti()
Dim i&
Dim j&

 Cells.Clear
 
    For i = 1 To 10
        For j = 1 To 10
            Cells(i, j) = i * j
            Cells(i, j).Font.Bold = True
            If i = 1 Or j = 1 Then Cells(i, j).Font.Color = -16776961
        Next
    Next
End Sub
