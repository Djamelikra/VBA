Attribute VB_Name = "mTrigonometry"
Option Explicit
'mTrigonometry

Sub arcTangente()
    Dim x&
    Dim convertToPi&
    'Pour convertir l’arctangente en degrés, multipliez le résultat par 180/PI( ) ou utilisez la fonction DEGRES.
'    convertToPi = 180 / Math.pi
    x = 1.33333333333 '= 4/3
    MsgBox Atn(x)
    
End Sub
Sub PI_as_Double_Value()
Dim pi As Double
pi = Application.WorksheetFunction.pi()
MsgBox pi
End Sub
