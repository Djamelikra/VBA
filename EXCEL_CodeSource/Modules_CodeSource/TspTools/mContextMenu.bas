Attribute VB_Name = "mContextMenu"
Option Explicit
'mContextMenu <=> ThisWorkbook
Dim Ch As Range

'"LAUNCH FUNCTIONS#########################################

Sub launch_functionModulo()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select the dividend X  ", "Scalar", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select the divisor  Y ", "Scalar", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = functionModulo(operandX, operandY)
    
    
End Sub

Sub launch_euclidianNorm2D()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select Matrix ordered distance X  ", "Scalar", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select Matrix ordered distance  Y ", "Scalar", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = euclidianNorm2D(operandX, operandY)
    
    
End Sub
Sub launch_AbsoluteDifference()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX#
    Dim operandY#
    Dim trackedX#
    Dim trackedY#

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select  R (L2) ", "Scalar", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select the mean (Mean Group) ", "Scalar", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = AbsoluteDifference(operandX, operandY)
    
    
End Sub
Sub launch_distanceSumSquare()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select Matrix ordered distance X ", "Scalar one", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select Matrix ordered distance Y ", "Scalar two", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = distanceSumSquare(operandX, operandY)
    
    
End Sub
Sub launch_rmRatio()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select R (L2) ", "Scalar one", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select the mean (Mean Group) ", "Scalar two", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = rmRatio(operandX, operandY)
    
    
End Sub
Sub launch_Surface()
    'Misaha
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance (masafa) 1 X: axis ", "Scalar one", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance (masafa) 2 Y: axis ", "Scalar two", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
    
    ActiveCell.Value = Surface(operandX, operandY)
    
    
End Sub
Sub launch_thetaAngleDeg()
    
    Dim Rng As Range
    Dim WorkRng As Range
    Dim operandX&
    Dim operandY&
    Dim trackedX&
    Dim trackedY&
    Dim thetaDegree# 'Mandatory
    

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance (masafa) X the abscissa axis ", "Scalar one", WorkRng.Address, Type:=8)
    operandX = WorkRng.Value

    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select distance (masafa) Y the ordinate axis ", "Scalar two", WorkRng.Address, Type:=8)
    operandY = WorkRng.Value
    
    
'to do format number for decimal after comma for angle in degree
'2023/08/19 12:00

    thetaDegree = thetaAngleDeg(cartesianToPolarTheta(operandX, operandY))
    

'    ActiveCell.Value = thetaAngleDeg(cartesianToPolarTheta(operandX, operandY)) & "°"  '=> No round & no format

'     ActiveCell.Value = Format(thetaDegree, "0.0") & "°" '=> No round

     
     ActiveCell.Value = Round(thetaDegree, 1) & "°"  '=> Round
     
    
    
End Sub
'"FUNCTIONS#########################################

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
Function functionModulo(ByVal X As Double, ByVal Y As Double) As Double
    'Modulo calculus
     functionModulo = X Mod Y
End Function
Function euclidianNorm2D(ByVal X As Double, ByVal Y As Double) As Double
    'The L² norm (Euclidian Norm) measures the shortest distance from the origin.
    'It is defined as the root of the sum of the squares of the components of the vector.
    'It can be calculated from the Cartesian coordinates (X,Y) of the points using the Pythagorean theorem.
     euclidianNorm2D = Sqr(X ^ 2 + Y ^ 2)
End Function
Function AbsoluteDifference(ByVal X As Double, ByVal Y As Double) As Double
    'Calculus The Absolute Difference between two numbers
     AbsoluteDifference = Abs(X - Y)
     
End Function
Function distanceSumSquare(ByVal X As Double, ByVal Y As Double) As Double
    'Calculus the sum of 2 distances squared
     distanceSumSquare = (X ^ 2 + Y ^ 2)
End Function
Function rmRatio(ByVal X As Single, ByVal Y As Single) As Single
'Function rmRatio(ByVal X As Double, ByVal Y As Double) As Double
    'Calculus the ratio of 2 scalars : R/Mean
    If Y = 0 Then
        MsgBox "division by zero impossible !", vbExclamation, "Tsp"
        Exit Function
    End If
     rmRatio = (X / Y)
End Function
Function Surface(ByVal X As Byte, ByVal Y As Byte) As Integer
    'Calculus the surface (misaha) of 2 scalars
     Surface = (X * Y)
End Function
Function thetaAngleDeg(tangAngle As Double) As Double
'returns the theta angle of a supplied number (tangente angle), in degree;
'3.141592653589793238562643383279
    Const pi As Double = 3.14159265358979
      'Pour convertir l’arctangente en degrés, il faut multiplier le résultat par 180/PI( )
    thetaAngleDeg = Atn(tangAngle) * 180 / pi
End Function
Function cartesianToPolarTheta(ByVal X As Double, ByVal Y As Double) As Double
'returns the ratio y/ x;
       If X = 0 Then
        MsgBox "division by zero impossible !", vbExclamation, "Tsp"
        Exit Function
    End If
    
    cartesianToPolarTheta = Y / X
End Function
Sub IsSymetric()
'to do
'launch_euclidianNorm2D
End Sub

