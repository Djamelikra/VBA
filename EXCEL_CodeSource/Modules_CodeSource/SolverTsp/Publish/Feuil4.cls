VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub FillFullMatrixWithOrdinalData()
'
' algo_1
'

'
    Range("B66:N78").Select
    Selection.Copy
    Range("B84").Select
    ActiveSheet.Paste
    Range("A97").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
     Call messagePerf
     
     
End Sub

Sub findAndLighTheValMin1()
'algo_2
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customGreen$
    
    customGreen = RGB(59, 255, 148)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value = 1 "
'    Set WorkRng = Application.Selection
Set WorkRng = Range("B103:N115")
    'Range("B66:Q81")
    
'    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customGreen
        End If
    Next
    
    
'    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub FindAndRemoveSupNumber()
'alg 03
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrRemoved%
      Dim remainingValues%
      
      'there are five values left
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrRemoved = 0
    For Each Rng In WorkRng
       
        If Rng.Value > CInt(trackedValue) Then
            Rng.Value = ""
        nbrRemoved = nbrRemoved + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrRemoved - [A89] 'less the diagonal
    
   
    MsgBox nbrRemoved & " items out of " & totalValues & " values have been removed." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
End Sub
Sub fillWithDataProcessed()
' algo_4
'
Dim sMsg As String
    sMsg = "Eurék@!"
    
Application.ScreenUpdating = False

    Range("B84:N96").Select
    Selection.Copy
    Range("B103").Select
    ActiveSheet.Paste
    Range("A116").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    
    
    ActiveWorkbook.Save
    
    Beep
'
       Call msgBoxTempMsgWait(sMsg, 3)
    
    
    
End Sub
Sub findAndLighValMinimumParam()
'alg_05
  'Color in orange a new value choosed as min
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrUpdated%
      Dim remainingValues%
      
        Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrUpdated = 0
    For Each Rng In WorkRng
       
        If Rng.Value = CInt(trackedValue) Then
          
        Rng.Interior.Color = orangeColor
        nbrUpdated = nbrUpdated + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrUpdated
   
    MsgBox nbrUpdated & " items out of " & totalValues & " values have been updated." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
    
End Sub
Sub findAndLighValMin()
'alg_06
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
Sub findAndLighValMinSecond()
'alg_06
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
                Rng.Interior.Color = RGB(139, 255, 191) 'light green
                Rng.Select
                'suited format for the Min:
                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub

Sub findAndLighValGradient()
'alg_07
 'Gradient
'
  'Color in orange a new value choosed as min
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrUpdated%
      Dim remainingValues%
      
        Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrUpdated = 0
    For Each Rng In WorkRng
       
        If Rng.Value > 0 And Rng.Value <= CInt(trackedValue) Then
          
        Rng.Interior.Color = orangeColor
        nbrUpdated = nbrUpdated + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrUpdated
   
    MsgBox nbrUpdated & " items out of " & totalValues & " values have been updated." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
                            
    Beep
    
                            
    If nbrUpdated < 2 Then
        MsgBox "=>DIRECT||==>", vbExclamation, "=>DIRECT||==>"
    
    End If
    Beep
    Beep
    
    
End Sub
Sub findAndLighValMax()
    'alg_08
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
                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub FillMatrixSimpleOrderNewTrialReset()
' algo_09
'
    Range("B66:N78").Select
    Selection.Copy
    Range("B103").Select
    ActiveSheet.Paste
    Range("A116").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
     MsgBox "OK!", vbInformation
     
     
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
Sub CellBackgdOrange()
'

     Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = orangeColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
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
Sub CellBackgRed()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub ClearCellsRowContents()
'
    Selection.ClearContents
    ActiveCell.Select
End Sub
Sub ClearAll()
'reset
    Selection.Clear
    ActiveCell.Select
End Sub
