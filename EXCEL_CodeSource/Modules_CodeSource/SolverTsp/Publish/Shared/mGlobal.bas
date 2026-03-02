Attribute VB_Name = "mGlobal"
Option Explicit
'1 048 576 lignes et 16 384 colonnes
Public Const gCR$ = "©Djamel CHABANE"
Public Const gAppName$ = "TSP SOLVER"


Sub usfMenuShow()
    usfMenu.Show
End Sub

Sub lastRow()
'    Dim endRow&
'    endRow = Range("A" & Rows.Count).End(xlUp).Row
'    MsgBox endRow, vbInformation
    Range("A1048576").Select
End Sub

 
Sub Macro1()
    MsgBox "Essai 01"
End Sub
 
 
Sub Macro2()
    MsgBox "Essai 02"
End Sub
