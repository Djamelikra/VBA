Attribute VB_Name = "mContextMenu"
Option Explicit
'mContextMenu <=> ThisWorkbook
Dim Ch As Range
Sub majuscule()
  For Each Ch In Selection
  'Pour chaque cellule Ch de ma sélection Selection
  'Selection permet de renvoyer l'objet sélectionné dans la fenêtre active
    If Not Ch.HasFormula Then
    'la propriété HasFormula de l'objet range permet de tester qu'aucune
    ' cellules de la plage ne contiennent de formule, dans ce cas la conversion est ignorée
      Ch.Value = UCase(Ch.Value) 'Upper Case = MAJUSCULE
    End If
  Next Ch
End Sub
'*************************************************************
Sub minuscule()
  For Each Ch In Selection
    If Not Ch.HasFormula Then
      Ch.Value = LCase(Ch.Value) 'Lower Case = minuscule
    End If
  Next Ch
End Sub

'*************************************************************
 Sub nompropre()
  For Each Ch In Selection
    If Not Ch.HasFormula Then
      Ch.Value = Application.Proper(Ch.Value)
      'si une fonction Excel n'a pas d'équivalent en VBA
      'utiliser le la fonction Excel (anglais) comme une méthode de l'objet application
    End If
  Next Ch
End Sub
'***************************************************************
Sub ClearCellContent()
'1088; 536
'Clear content
    Selection.ClearContents
    ActiveCell.Select
End Sub
Sub launch_euclidianNorm2D()
    
    Dim rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select operand X ", "Scalar", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select operand Y ", "Scalar", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = euclidianNorm2D(operandX, operandY)
    
    
End Sub
Sub launch_distanceSumSquare()
    
    Dim rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance X ", "Scalar one", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance Y ", "Scalar two", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = distanceSumSquare(operandX, operandY)
    
    
End Sub

Function euclidianNorm2D(ByVal X As Double, ByVal Y As Double) As Double
    'The L² norm (Euclidian Norm) measures the shortest distance from the origin.
    'It is defined as the root of the sum of the squares of the components of the vector.
    'It can be calculated from the Cartesian coordinates (X,Y) of the points using the Pythagorean theorem.
     euclidianNorm2D = Sqr(X ^ 2 + Y ^ 2)
End Function
Function distanceSumSquare(ByVal X As Double, ByVal Y As Double) As Double
    'Calculus the sum of 2 distances squared
     distanceSumSquare = (X ^ 2 + Y ^ 2)
End Function
Sub IsSymetric()
'to do
'launch_euclidianNorm2D
End Sub

