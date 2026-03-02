VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfSearchValue 
   Caption         =   "Search value"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   OleObjectBlob   =   "usfSearchValue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfSearchValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'param
     With cboOperator
     'param
        .Style = fmStyleDropDownList
        
     'fill cbo
        .AddItem (">")
        .AddItem ("<")
        .AddItem (">=")
        .AddItem ("=<")
        .AddItem ("<>")
        .AddItem ("=")
        
         'param
        .ListIndex = 0
     End With
    
End Sub

Private Sub cmdOK_Click()
    Dim addRange$
    Dim oRng, oCell As Range
    Dim minVal#
    Dim intResult&
    Dim mathOperator$
    Dim mathOperand&
    
  
    
    addRange = refSelection.Value
    Set oRng = Range(addRange)
    
    
    
'    minVal = WorksheetFunction.Min(oRng)
    If cboOperator.ListIndex = -1 Or refOperand.Value = "" Then
        Exit Sub
'        cboOperator.SetFocus
    Else
        mathOperator = cboOperator.Value
        mathOperand = Range(refOperand.Value).Value
    End If
    
    
    
    For Each oCell In oRng
        If oCell.Value > mathOperand Then
'             Dim intResult As Long
           'displays the color dialog
            Application.Dialogs(xlDialogEditColor).Show 40, 100, 100, 200
            'gets the color selected by the user
            intResult = ThisWorkbook.Colors(40)
            'changes the fill color of cell A1
'            Range("A1").Interior.Color = intResult
            oCell.Interior.Color = intResult
        End If
    Next oCell
    
'    Set workRng = Application.Selection
'Set workRng = Application.InputBox("Range", xTitleId, workRng.Address, Type:=8)
'For Each rng In workRng
'    If rng.Value > 10 Then
''        Rng.Value = ""
'         rng.Interior.Color = vbYellow
'    End If
    
    
    
End Sub
Private Sub cmdColor_Click()
    Application.Dialogs(xlDialogEditColor).Show 40, 100, 100, 200
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub









