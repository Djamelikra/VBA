VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Const DATA_RANGE As String = "Select the full Data Range"
Private Const PERF As String = "Performed"
'Range("B66:Q81")
'template


Sub FillFullMatrixWithOrdinalData()
'
' FillFullMatrixWithOrdinalData Macro
'

'
 Range("B24:Q39").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=20
    Range("B45").Select
    ActiveSheet.Paste
    Range("A61").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save

       Call messagePerf


End Sub
Sub findAndLighTheValMin1()
'02
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
Set WorkRng = Range("B66:Q81")
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
Sub findAndLighTheValMin2()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(167, 255, 207)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value = 2"
'    Set WorkRng = Application.Selection
    Set WorkRng = Range("B66:Q81")
'    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 2
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
'    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighTheValMin3()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(255, 255, 91)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value = 3"
'    Set WorkRng = Application.Selection
    Set WorkRng = Range("B66:Q81")
'    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 3
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
'    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighTheValMin4()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(255, 212, 125)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighTheValMin5()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(255, 192, 0)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighTheValMin6()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(234, 139, 0)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighTheValMin7()
'0
    'Update cell in
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customColor$
    
    customColor = RGB(91, 212, 255)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customColor
        End If
    Next
    
    
    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub

Sub messagePerf()
    Beep
    Call msgBoxTempMsgWait(UCase$(PERF), 2)
    
End Sub
Sub msgBoxTempMsgWait(sMessage As String, secondWait As Byte)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after ? seconds
 
    Select Case InfoBox.Popup(sMessage, secondWait, gAppName, vbOKOnly)
        Case 1, -1
            Exit Sub
    End Select
    
    Beep
    
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
    Set WorkRng = Application.InputBox(DATA_RANGE, , WorkRng.Address, Type:=8)
    
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
'01
' fillWithDataProcessed Macro
'

  Dim sMsg As String
    sMsg = "Eurék@!"
    
    
  Application.ScreenUpdating = False
  
    ActiveWindow.SmallScroll Down:=-4
    Range("B45:Q60").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=13
    Range("B66").Select
    ActiveSheet.Paste
    Range("A82").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
      Application.ScreenUpdating = True
      
      Beep
      
       Call msgBoxTempMsgWait(sMsg, 3)
      
End Sub
Sub findAndLighValMinimumParam()
'05
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
    Set WorkRng = Application.InputBox(DATA_RANGE, , WorkRng.Address, Type:=8)
    
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
    Set WorkRng = Application.InputBox(DATA_RANGE, , WorkRng.Address, Type:=8)
    
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

Sub FillMatrixSimpleOrder()
'
' FillMatrixSimpleOrder Macro
'

'
    Range("U24:AJ39").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=37
    Range("B66").Select
    ActiveSheet.Paste
    Range("A82").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
    MsgBox "OK!", vbInformation
    
End Sub
Sub ClearCellsRow()
'
    Selection.ClearContents
    ActiveCell.Select
End Sub
Sub ClearAll()
'reset
    Selection.Clear
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

Sub CellBackgdGray()
'

     Dim GrayColor$
    
    GrayColor = RGB(128, 128, 128)
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = GrayColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub

Sub newTrial()

Call fillWithDataProcessed '01
Call findAndLighTheValMin1 '02
Call findAndLighTheValMin2
Call findAndLighTheValMin3



End Sub
'###################################################################To ADD useful#####################################

Private Sub Mandatory()

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
