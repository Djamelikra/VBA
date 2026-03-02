Attribute VB_Name = "mUsual"
Option Explicit

Sub msgBoxTemp(secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup("Click OK (this window closes automatically after)." & " " & secondWait & " " & "seconds.", _
    secondWait, gAppName, vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Sub CellBackgdYellow()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
