Attribute VB_Name = "mMaxMin"
Option Explicit
Sub minValueInRange()
'Determines the minimum value in a selectioned range
    Dim oRng As Range
    Dim valMin&
    
    Set oRng = Selection
    
    valMin = Application.WorksheetFunction.Min(oRng)
    
    MsgBox valMin, vbInformation, gAppName & " " & gCR
    
End Sub

Sub manValueInRange()
'Determines the maximum value in a selectioned range
    Dim oRng As Range
    Dim valMin&
    
    Set oRng = Selection
    
    valMin = Application.WorksheetFunction.Max(oRng)
    
    MsgBox valMin, vbInformation, gAppName & " " & gCR
    
End Sub
Sub findAndLighted()
'Updateby Extendoffice
Dim rng As Range
Dim WorkRng As Range
Dim xTitleId$
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each rng In WorkRng
    If rng.Value > 10 Then
'        Rng.Value = ""
         rng.Interior.Color = vbYellow
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
Sub findAndLighValMin()
    'Update cell in yellow
    Dim rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin&
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each rng In WorkRng
        If rng.Value = valMin Then
            If rng.Interior.Color = vbYellow Then
                rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                rng.Interior.Color = vbYellow
        End If
    Next
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValParameter()

Dim rng As Range
Dim WorkRng As Range
  Dim xTitleId$
   Dim valParameter&
   
On Error Resume Next
xTitleId = "Looking with Parameter value"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)

  valParameter = Application.InputBox("Enter the Parameter value", "Parameter", 0, 1)
For Each rng In WorkRng
    If rng.Value > valParameter Then
        rng.Interior.Color = vbYellow
    End If
Next
    
  
     
End Sub

