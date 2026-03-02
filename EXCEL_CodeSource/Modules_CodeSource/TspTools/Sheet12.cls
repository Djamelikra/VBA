VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub FindAndReplaceNumberBy1()
'to be adpated
    Dim rng As Range
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
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrRemoved = 0
    For Each rng In WorkRng
       
        If rng.Value = CInt(trackedValue) Then
            rng.Value = 1
        nbrRemoved = nbrRemoved + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrRemoved
    
   
    MsgBox nbrRemoved & " items out of " & totalValues & " values have been removed." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
End Sub

Sub FindAndRemoveSupNumber()
'to be adpated
    Dim rng As Range
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
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrRemoved = 0
    For Each rng In WorkRng
       
        If rng.Value > CInt(trackedValue) Then
            rng.Value = ""
        nbrRemoved = nbrRemoved + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrRemoved
    
   
    MsgBox nbrRemoved & " items out of " & totalValues & " values have been removed." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
End Sub

Sub findAndLighValMin()
    'Update cell in yellow
    Dim rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin!
      
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
Sub FillDataFullMatrixTsp14()
'
' FillDataHalfMatrixTsp14 Macro
'

  Application.ScreenUpdating = False
  
    Range("B77:O90").Select
    Selection.Copy
    Range("B97").Select
    ActiveSheet.Paste
    Range("A111").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
      Application.ScreenUpdating = True
    Beep
    
    Call msgBoxTemp(3)
    
    
'    MsgBox "Eurék@!", vbInformation, gAppName & " " & gCR
    
End Sub
Sub findAndLighValMinimum1()
    'Update cell in yellow
    Dim rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customGreen$
    
    customGreen = RGB(59, 255, 148)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each rng In WorkRng
        If rng.Value = valMin Then
            If rng.Interior.Color = vbYellow Then
                rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                rng.Interior.Color = customGreen
        End If
    Next
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValMinimumParam()
  'Color in orange a new value choose as min
    Dim rng As Range
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
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrUpdated = 0
    For Each rng In WorkRng
       
        If rng.Value = CInt(trackedValue) Then
          
        rng.Interior.Color = orangeColor
        nbrUpdated = nbrUpdated + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrUpdated
   
    MsgBox nbrUpdated & " items out of " & totalValues & " values have been updated." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
    
End Sub


Sub FillMatriNearx()
'
' FillMatriNearx Macro
'

'
    Range("Y63:AL76").Select
    Selection.Copy
    Range("B77").Select
    ActiveSheet.Paste
    Range("A91").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
End Sub



















