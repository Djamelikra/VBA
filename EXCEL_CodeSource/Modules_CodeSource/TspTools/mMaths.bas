Attribute VB_Name = "mMaths"
Option Explicit
'mMaths
Function MeanOpt2Args5(arg1 As Single, arg2 As Single, Optional arg3 As Single, Optional arg4 As Single, Optional arg5 As Single) As Single
'Calculus arithmetic mean with 2 mandatory arguments
'     Application.Volatile 'switch
     
     If IsMissing(arg3) And IsMissing(arg4) And IsMissing(arg5) Then
        MeanOpt2Args5 = (arg1 + arg2) / 2
     
      ElseIf IsMissing(arg4) And IsMissing(arg5) Then
        MeanOpt2Args5 = (arg1 + arg2 + arg3) / 3
     
        ElseIf IsMissing(arg5) Then
        MeanOpt2Args5 = (arg1 + arg2 + arg3 + arg4) / 4
        
        Else
        MeanOpt2Args5 = (arg1 + arg2 + arg3 + arg4 + arg5) / 5
     End If

End Function

Sub Mean2to5()
      Dim Rng As Range
    Dim WorkRng As Range
      Dim arg1!, arg2!, arg3!, arg4!, arg5!
      
        Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
    
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select arg1 value", "One cell", WorkRng.Address, Type:=8)
    arg1 = WorkRng.Value
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select arg2 value", "One cell", WorkRng.Address, Type:=8)
    arg2 = WorkRng.Value
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select arg3 value", "One cell", WorkRng.Address, Type:=8)
    arg3 = WorkRng.Value
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select arg4 value", "One cell", WorkRng.Address, Type:=8)
    arg4 = WorkRng.Value
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select arg5 value", "One cell", WorkRng.Address, Type:=8)
    arg5 = WorkRng.Value
    
    'calculus:
    
    MsgBox MeanOpt2Args5(arg1, arg2, arg3, arg4, arg5), vbInformation, "Mean"
    
End Sub
Sub findAndLighValArgs()
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
    
                            
    If totalValues < 2 Then
        MsgBox "=>DIRECT||==>", vbExclamation, "=>DIRECT||==>"
    
    End If
    Beep
    Beep
    
    
End Sub

Public Sub Mandatory()

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
Function arithMean() As Single
'to do

End Function
Function arithMeanAbsDifference() As Single
'MAD
'to do

End Function
Function Surface()
'Surface
'to do

End Function


Sub boiteDialogue(nom As String, Optional prenom, Optional age)

  If IsMissing(prenom) Then 'Si le prénom est manquant, on n'affiche que le nom
            MsgBox nom
        Else 'Sinon, on affiche le nom et le prénom
            MsgBox nom & " " & prenom
        End If
        
'    MsgBox nom
'    MsgBox prenom
End Sub


Sub testBD()
    Dim nom As String, prenom As String   ', age As Integer
     nom = "Ranxx"
    prenom = "Ranbbb"
    
    boiteDialogue nom, prenom
    
End Sub
