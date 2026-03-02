VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfDirac 
   Caption         =   "Dirac"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   OleObjectBlob   =   "usfDirac.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfDirac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
'#### Pass Full Array to ComboBox
    
'    Dim pointsArray(13, 1) As Variant
    
    With cboDirac
        .Width = 125
        .ColumnCount = 2
        .ColumnWidths = "20;0"
        '#2020/01/16 17:08-> Dir/ HCP
        .BoundColumn = 2
        '#2020/01/16 17:08-> Dir/ HCP
    End With
    
    'populate array
'    pointsArray(0, 0) = "1"
'    pointsArray(0, 1) = "DIR_1"
'    pointsArray(1, 0) = "2"
'    pointsArray(1, 1) = "DIR_2"
'    pointsArray(2, 0) = "3"
'    pointsArray(2, 1) = "DIR_3"
'
'    pointsArray(3, 0) = "4"
'    pointsArray(3, 1) = "DIR_4"
'
'    pointsArray(4, 0) = "5"
'    pointsArray(4, 1) = "DIR_5"
'
'    pointsArray(5, 0) = "6"
'    pointsArray(5, 1) = "DIR_6"
    
'    pointsArray(3, 0) = "7"
'    pointsArray(3, 1) = "DIR_7"
'    pointsArray(3, 0) = "8"
'    pointsArray(3, 1) = "DIR_8"
'    pointsArray(3, 0) = "9"
'    pointsArray(3, 1) = "DIR_9"
'    pointsArray(3, 0) = "10"
'    pointsArray(3, 1) = "DIR_10"
'    pointsArray(3, 0) = "11"
'    pointsArray(3, 1) = "DIR_11"
'    pointsArray(3, 0) = "12"
'    pointsArray(3, 1) = "DIR_12"
'    pointsArray(3, 0) = "13"
'    pointsArray(3, 1) = "DIR_13"
    
    'then populate combobox with full array
'    cboDirac.List = pointsArray
    '######################
    Dim i As Integer
    Dim tabPoints(1 To 13, 1 To 2) As Variant
    For i = 1 To 13
        tabPoints(i, 1) = i
        tabPoints(i, 2) = "DIR_" & i
    Next i
    cboDirac.List = tabPoints
    
    If IsNull(cboDirac) Then
        cmdRunDir.Visible = False
    End If
 
End Sub

Private Sub cboDirac_Change()

    With cboDirac
        .Width = 70
    End With
    
    cmdRunDir.Visible = True
    
    cmdRunDir.SetFocus
    
End Sub

Private Sub cmdRunDir_Click()
     
     Run Me.cboDirac.Value
     
     Unload Me
     
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub
