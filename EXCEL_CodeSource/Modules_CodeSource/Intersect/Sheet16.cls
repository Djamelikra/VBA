VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit
'

Sub CopyDataDirac_N_6_SumCoef1()
'
' CopyDataDirac_N_6_SumCoef1 Macro
    Range("C96:O108").Select
    Selection.Copy
    Range("C120").Select
    ActiveSheet.Paste
    Range("B133").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
End Sub


'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    If Selection.Count = 1 Then
'
'                If Not Intersect(Target, [vMinP1]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP1
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP2]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP2
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP3]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP3
''                    Exit Sub
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP4]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP4
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                    If Not Intersect(Target, [vMinP5]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP5
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP6]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP6
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP7]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP7
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP8]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP8
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                    If Not Intersect(Target, [vMinP9]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP9
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP10]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP10
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP11]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP11
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP12]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP12
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'                If Not Intersect(Target, [vMinP13]) Is Nothing Then 'we take only the case where is a cell to be clicked
'                    Call delDatavMinP13
''                            MsgBox "In the data ", vbInformation, gAppname
'                End If
'
'
'    End If
'End Sub
