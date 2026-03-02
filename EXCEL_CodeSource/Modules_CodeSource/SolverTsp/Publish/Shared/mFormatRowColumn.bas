Attribute VB_Name = "mFormatRowColumn"
Option Explicit

Sub HideColumn()
Attribute HideColumn.VB_Description = "Hide a column"
Attribute HideColumn.VB_ProcData.VB_Invoke_Func = " \n14"
'
' HideColumn Macro
' Hide a column
'

'
'    Columns("U:U").Select
'    Range("U42").Activate
    Selection.EntireColumn.Hidden = True
End Sub
Sub HideRow()
Attribute HideRow.VB_Description = "Hide a Row"
Attribute HideRow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' HideRow Macro
' Hide a Row
'

'
'    Rows("51:51").Select
'    Range("B51").Activate
    Selection.EntireRow.Hidden = True
End Sub
Sub ColumnsFit11cm()
'

    Selection.ColumnWidth = 11
    ActiveCell.Select
End Sub
Sub ColumnsFit07cm()
'

    Selection.ColumnWidth = 7
    ActiveCell.Select
End Sub
Sub UnhideAll()
'
' UnhideAll Macro
' Unhide data of columns and rows
'

'
    Cells.Select
    Range("A1").Activate
    Selection.EntireRow.Hidden = False
    Selection.EntireColumn.Hidden = False
'    ActiveWindow.SmallScroll Down:=-41
    Range("A1").Select
End Sub
