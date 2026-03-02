Attribute VB_Name = "mSearchValue"
Option Explicit


Sub findAndLightvalMaxd()
'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim trackedValue&
    
    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
    'On Error Resume Next
    xTitleId = "Cell lighted"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    For Each Rng In WorkRng
        If Rng.Value > CInt(trackedValue) Then
    '        Rng.Value = ""
             Rng.Interior.Color = vbYellow
        End If
    Next
End Sub
Sub FindReplace()
    'to be adpated
    Dim Rng As Range
    Dim WorkRng As Range
       Dim trackedValue&
    
    trackedValue = Application.InputBox("Searched value:", "Tracking")
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    For Each Rng In WorkRng
        If Rng.Value > CInt(trackedValue) Then
            Rng.Value = 0
        End If
    Next
End Sub
Sub FindRemove()
'to be adpated
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
    
    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
    For Each Rng In WorkRng
        If Rng.Value > CInt(trackedValue) Then
            Rng.Value = ""
        End If
    Next
End Sub
