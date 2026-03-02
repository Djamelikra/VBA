VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfTSP 
   Caption         =   "TSP"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4650
   OleObjectBlob   =   "usfTSP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfTSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oTool As cTool

Private Sub ConstructorTSP(boundTypeData As Integer)
    'To populate cboTSP
'#### Pass Full Array to ComboBox

    '#Usf
    With usfTSP
        .StartUpPosition = 3
'        .Left = Application.Width - .Width
        .Width = 225
        .Height = 283
        .Left = 765
        .Top = 3
    End With

    With cboTSP
        .Width = 40
        .ColumnCount = 5
        .ColumnWidths = "20;0;0;0;0"
        '#2020/01/16 17:08-> Dir/ HCP
        .BoundColumn = boundTypeData
        .ListRows = 13
        '#2020/01/16 17:08-> Dir/ HCP
    End With
    
    '######################
    Dim i As Integer
    Dim tabPoints(1 To 13, 1 To 6) As Variant
    For i = 1 To 13
        tabPoints(i, 1) = i
        tabPoints(i, 2) = "DirY_" & i
        tabPoints(i, 3) = "DirXY_" & i
        tabPoints(i, 4) = "Lim_" & i
        tabPoints(i, 5) = "ClearXY_" & i
        tabPoints(i, 6) = "ClearX_" & i
    Next i
    cboTSP.List = tabPoints
    
    If IsNull(cboTSP) Then
        cmdRunDir.Visible = False
        cmdClearYX.Visible = False
        cmdHightLimit.Visible = False
        cmdRunDirXY.Visible = False
        cmdClearX.Visible = False
    End If
    
End Sub









Private Sub cmdClose_Click()
    Beep
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
'© Djamel CHABANE 25/01/2020 16:40
'By dafault on Dir
    Call ConstructorTSP(2)
    
End Sub

Private Sub cboTSP_AfterUpdate()

    With cboTSP
        .Width = 40
    End With
    
       cmdRunDir.Visible = True
        cmdClearYX.Visible = True
        cmdHightLimit.Visible = True
        cmdRunDirXY.Visible = True
        cmdClearX.Visible = True

    cmdRunDir.SetFocus
    
End Sub


Private Sub cmdRunDir_Click()
     With Me.cboTSP
        .BoundColumn = 2
        Run .Value
     End With
'    Run Me.cboTSP.Value
'    Unload Me
End Sub
Private Sub cmdRunDirXY_Click()
'© Djamel CHABANE 26/01/2020 18:15
  With Me.cboTSP
        .BoundColumn = 3
        Run .Value
     End With
End Sub
Private Sub cmdHightLimit_Click()
'© Djamel CHABANE 25/01/2020 16:40
  With Me.cboTSP
        .BoundColumn = 4
        Run .Value
     End With

End Sub

Private Sub cmdClearYX_Click()
 Set oTool = New cTool
     With Me.cboTSP
        .BoundColumn = 5
        Run .Value
     End With
      With oTool
        .MsgPopup
     End With
'    Unload Me
End Sub

Private Sub cmdClearX_Click()
    Set oTool = New cTool
    
   With Me.cboTSP
        .BoundColumn = 6
        Run .Value
     End With
     With oTool
        .MsgPopup
     End With
End Sub

',###############################################################################################

'© Djamel CHABANE 28/01/2020 10:37
Private Sub cmdFraction_Click()
    Call Fraction
End Sub
Private Sub cmdNbrStd_Click()
    Call NumberStd
End Sub

Private Sub Fraction()
' Fraction
    Selection.NumberFormat = "# ?/?"
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Private Sub NumberStd()
'
'    Selection.NumberFormat = "General"
    Selection.NumberFormat = "0.00"
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Private Sub cmdSetOrdre0_Click()
    Call SetTsp13Ordre0
End Sub
Private Sub SetTsp13Ordre0()
    Range("B21:N33").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A53").Select
End Sub

Private Sub cmdSetCellDirac_Click()
    Call DelCellSupToDir
End Sub

Private Sub DelCellSupToDir()


Dim ObjCell As Range

For Each ObjCell In Range("B37:N49").Cells
    If ObjCell.Value > 6 Then
        ObjCell.ClearContents
    End If
Next
End Sub
Private Sub cmdMandatory_Click()
    Call Mandatory
End Sub
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
End Sub
Private Sub cmdRDN_Click()
    Call RDN
    
End Sub
Private Sub RDN()
'Regular Dirac Normalization

    Range("R21:AD33").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A36").Select
End Sub
Private Sub DCN()
'Dirac's Chiral Normalization

    Range("AI12:AU24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    
    
    Range("B37:N49").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("A48").Select
End Sub
Private Sub cmdDCN_Click()
    Call DCN
End Sub

Private Sub HG()
'Hamiltonian graph
'© Djamel CHABANE 04/02/2020 18:31
    Range("BF12:BR24").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A46").Select
End Sub
Private Sub cmdHG_Click()
'© Djamel CHABANE 04/02/2020 18:31
    Call HG

End Sub
Sub RDN_Graph()
'© Djamel CHABANE 04/02/2020 21:41
    Range("BF30:BR42").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A49").Select
End Sub
Private Sub cmdRDNGraph_Click()
'© Djamel CHABANE 04/02/2020 21:41
    Call RDN_Graph

End Sub
Sub DCN_O2()
'© Djamel CHABANE 05/02/2020 15:00
    Range("AH37:AT49").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A50").Select
End Sub
Private Sub cmdDCNO2_Click()
'© Djamel CHABANE 05/02/2020 15:00

    Call DCN_O2
    
End Sub
Sub RDN_HG()
'© Djamel CHABANE 06/02/2020 18:41
    Range("BV30:CH42").Select
    Selection.Copy
    Range("B37").Select
    ActiveSheet.Paste
    Range("A50").Select
End Sub
