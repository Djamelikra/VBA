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

Sub maxValueInRange()
'Determines the maximum value in a selectioned range
    Dim oRng As Range
    Dim valMin&
    
    Set oRng = Selection
    
    valMin = Application.WorksheetFunction.Max(oRng)
    
    MsgBox valMin, vbInformation, gAppName & " " & gCR
    
End Sub
Sub findAndLighted()
'Updateby Extendoffice
Dim Rng As Range
Dim WorkRng As Range
Dim xTitleId$
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    If Rng.Value > 10 Then
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
Sub findAndLighValMin()
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(255, 192, 0)
                Rng.Select
                'suited format for the Min:
                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValMax()
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMax!
      
    On Error Resume Next
    xTitleId = "Looking for MAX value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMax = Application.WorksheetFunction.Max(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMax Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(255, 192, 0)
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub

Sub findAndLighValMinOverZero()
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
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(255, 192, 0)
        End If
    Next
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub

Sub findAndLighValParameter()

Dim Rng As Range
Dim WorkRng As Range
  Dim xTitleId$
   Dim valParameter&
   
On Error Resume Next
xTitleId = "Looking with Parameter value"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)

  valParameter = Application.InputBox("Enter the Parameter value", "Parameter", 0, 1)
For Each Rng In WorkRng
    If Rng.Value > valParameter Then
        Rng.Interior.Color = vbYellow
    End If
Next
    
  
     
End Sub


Sub Mandatory()

    With Selection.Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(153, 229, 255)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

