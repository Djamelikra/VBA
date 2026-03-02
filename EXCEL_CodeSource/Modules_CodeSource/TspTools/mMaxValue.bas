Attribute VB_Name = "mMaxValue"
Option Explicit
'mMaxValue
Sub test_maxOf2Numbers()
'Determines the maximum of 2 numbers
   MsgBox maxOf2Numbers(13, 23), vbInformation, gAppName
End Sub

Function maxOf2Numbers(dblX As Double, dblY As Double) As Double
    'Determines the maximum of 2 numbers
    maxOf2Numbers = Application.WorksheetFunction.Max(dblX, dblY)
End Function
