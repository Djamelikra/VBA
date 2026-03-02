VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub FindAndRemoveSupNumberSingle()
'to be adpated
    Dim rng As Range
    Dim WorkRng As Range
      Dim trackedValue!
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    For Each rng In WorkRng
        If rng.Value > CInt(trackedValue) Then
            rng.Value = ""
        End If
    Next
End Sub

Sub findAndLighValMinSingle()
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
                Beep
                rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
            'tada
            Call tadaSound
                rng.Interior.Color = vbYellow
        End If
    Next
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub FillDataHalfMatrixTsp14_Single()
'
' FillDataHalfMatrixTsp14_Single Macro
'
Application.ScreenUpdating = False
'
    Range("B77:O90").Select
    Selection.Copy
    Range("B97").Select
    ActiveSheet.Paste
    Range("A114").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
    
    Application.ScreenUpdating = True
    
    Call msgBoxTemp(3)
    

    
End Sub
Sub copyDiagoal()
'
' copyDiagoal Macro
'

'
    Range("B97").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("C98").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("D99").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("E100").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("F101").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("G102").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("H103").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("I104").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("J105").Select
    ActiveCell.FormulaR1C1 = "9"
    Range("K106").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("L107").Select
    ActiveCell.FormulaR1C1 = "11"
    Range("M108").Select
    ActiveCell.FormulaR1C1 = "12"
    Range("N109").Select
    ActiveCell.FormulaR1C1 = "13"
    Range("O110").Select
    ActiveCell.FormulaR1C1 = "14"
    Range("A115").Select
    ActiveWorkbook.Save
End Sub

