Attribute VB_Name = "mMaxMin"
Option Explicit
Sub minValueInRange()
'Determines the minimum value in a selectioned range
    Dim oRng As Range
    Dim valMin&
    
    Set oRng = Selection
    
    valMin = Application.WorksheetFunction.Min(oRng)
    
    MsgBox valMin, vbInformation, gAppName & " " & gCr
    '
    
    
End Sub

Sub maxValueInRange()
'Determines the maximum value in a selectioned range
    Dim oRng As Range
    Dim valMax&
    
    Set oRng = Selection
    
    valMax = Application.WorksheetFunction.Max(oRng)
    
    MsgBox valMax, vbInformation, gAppName & " " & gCr
    
End Sub
Sub findAndLightvalMaxd()
'Update cell in yellow
Dim Rng As Range
Dim WorkRng As Range
Dim xTitleId$
On Error Resume Next
xTitleId = "Cell lighted"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    If Rng.Value > 10 Then
'        Rng.Value = ""
         Rng.Interior.Color = vbYellow
    End If
Next
End Sub
Sub findAndLightvalMin()
'Update cell in yellow
Dim Rng As Range
Dim WorkRng As Range
Dim xTitleId$
Dim valMin&
  
On Error Resume Next
xTitleId = "Looking for min value"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)

valMin = Application.WorksheetFunction.Min(WorkRng)

For Each Rng In WorkRng
    If Rng.Value = valMin Then
'        Rng.Value = ""
         Rng.Interior.Color = vbYellow
    End If
Next
End Sub
Sub greatValue()
    Dim i As Integer
    Dim Cells_Value As String
    For i = 5 To 12
        If Cells(i, 3).Value > 10 Then
            Cells_Value = Cells(i, 3).Value
            MsgBox "Value is " & Cells_Value
            Exit For
        End If
    Next
End Sub
