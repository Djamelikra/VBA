Attribute VB_Name = "mTspFunctions"
Option Explicit
'mTspFunctions
Function reductionNormalizationTo10(Numerator As Integer, Denominator As Integer) As Single
Attribute reductionNormalizationTo10.VB_Description = "Reduces and normalizes the dataset from the distance matrix to 10"
'Reduces and normalizes the dataset from the distance matrix to 10
    'Numerator = Min
    'Denominator = MaxMin
    'TO10 = Max(MaxMin) : according to the distance matrix dataset
    Const TO10 As Integer = 10
    
    reductionNormalizationTo10 = (Numerator / Denominator) * TO10
    
End Function

Sub test_reductionNormalizationTo10()

    MsgBox reductionNormalizationTo10(1, 100), vbInformation
    

End Sub
