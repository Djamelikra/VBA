Attribute VB_Name = "mDistance"
Option Explicit

Function hypotenuse(co As Double, ca As Double) As Double
Attribute hypotenuse.VB_Description = "hypotenuse calculation"
    'Calcul de l'hypoténuse d'un triangle rectangle
    'co = coté opposé; ca = coté adjacent
    hypotenuse = Sqr(co ^ 2 + ca ^ 2)
End Function
Function eucliDistance2D(oppS As Double, adjS As Double) As Double
Attribute eucliDistance2D.VB_Description = "Euclidean distance between two points in Euclidean space is the length of a line segment between the two points. It can be calculated from the Cartesian coordinates of the points using the Pythagorean theorem. oppS = opposite side; adjS = adjacent side."
    'Euclidean distance between two points in Euclidean space is the length of a line segment between the two points.
    'It can be calculated from the Cartesian coordinates of the points using the Pythagorean theorem.
    'oppS = opposite side; adjS = adjacent side
    eucliDistance2D = Sqr(oppS ^ 2 + adjS ^ 2)
End Function
Function euclidianNorm(X As Double, Y As Double) As Double
Attribute euclidianNorm.VB_Description = "The L² norm (Euclidian Norm) measures the shortest distance from the origin.It is defined as the root of the sum of the squares of the components of the vector.It can be calculated from theCartesian coordinates (X,Y) of the points using the Pythagorean theorem."
    'The L² norm (Euclidian Norm) measures the shortest distance from the origin.
    'It is defined as the root of the sum of the squares of the components of the vector.
    'It can be calculated from the Cartesian coordinates (X,Y) of the points using the Pythagorean theorem.
    euclidianNorm = Sqr(X ^ 2 + Y ^ 2)
End Function

 
 Sub launch_functionDialog()
    Call SendKeys("+{f3}", True)
    
'    ActiveCell.Formula2R1C1 = "=euclidianNorm()"
End Sub
Function euclidianNorm2D(ByVal X As Double, ByVal Y As Double) As Double
    'The L² norm (Euclidian Norm) measures the shortest distance from the origin.
    'It is defined as the root of the sum of the squares of the components of the vector.
    'It can be calculated from the Cartesian coordinates (X,Y) of the points using the Pythagorean theorem.
     euclidianNorm2D = Sqr(X ^ 2 + Y ^ 2)
End Function

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

Sub launch_notepad()
    Call Shell("C:\windows\system32\Notepad.exe", vbNormalFocus)
    Application.Wait (Now() + TimeValue("00:00:3"))
   Call SendKeys("{F5} {Enter} Hello !", True)
End Sub
Sub testHyp()
    MsgBox hypotenuse(5, 4)
    
End Sub
