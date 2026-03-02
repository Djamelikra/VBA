VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()

'2023_05_14 18:35
'#Context Menu
' Menu accessible quand on fait le clique avec le bouton de droit dans une cellule

 
'    On Error Resume Next
    'On efface le menu
'    Application.CommandBars("Cell").Reset
    
    Call resetMenu
    
    '#Police case
       Dim Cpop1 As CommandBarPopup
    Dim Cbut As CommandBarButton
    Dim MaBarre As CommandBar
    Set MaBarre = Application.CommandBars("cell")
    Set Cpop1 = MaBarre.Controls.Add(msoControlPopup)
    'on ajoute le popup qui va comporté tout les boutons
    With Cpop1
        .Caption = "Modifier la casse"    'label du sous-menu
    End With
    'on ajoute un bouton au sous-menu contextuel
    Set Cbut = Cpop1.Controls.Add(Type:=msoControlButton)
    With Cbut
        .FaceId = 403    ' bouton avec icône (référence à une table)
        .Caption = "MAJUSCLE"    ' label du bouton
        .OnAction = "majuscule"    'appel de la macro associée
    End With
    'on répéte l'opération pour les deux autres boutons
    Set Cbut = Cpop1.Controls.Add(Type:=msoControlButton)
    With Cbut
        .FaceId = 404
        .Caption = "minuscule"
        .OnAction = "minuscule"
    End With
    Set Cbut = Cpop1.Controls.Add(Type:=msoControlButton)
    With Cbut
        .FaceId = 306
        .Caption = "Nom Propre"
        .OnAction = "nompropre"
    End With
    '/#Police case
    
    
'    Dim Cbut As CommandBarButton
'#launch_functionDialog
    Set Cbut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
 
    'On le remet
    With Cbut
       .BeginGroup = False
       .Caption = "Functions..."
       .Style = msoButtonIconAndCaption
       .FaceId = 385
       .OnAction = "launch_functionDialog"
    End With
    
    '#launch_euclidianNorm2D
       Set Cbut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
    With Cbut
       .BeginGroup = False
       .Caption = "Euclidian Norme L²"
       .Style = msoButtonIconAndCaption
       .FaceId = 521
       .OnAction = "launch_euclidianNorm2D"
    End With
    
       '#launch_distanceSumSquare
       Set Cbut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
    With Cbut
       .BeginGroup = False
       .Caption = "Distances Sum Square"
       .Style = msoButtonIconAndCaption
       .FaceId = 966 '1042 '378  '521
       .OnAction = "launch_distanceSumSquare"
    End With
    
    
    '#clear cells
    Set Cbut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
    With Cbut
        .BeginGroup = False
        .Style = msoButtonIconAndCaption
        .FaceId = 536   '1088
        .Caption = "Delete Content"
        .OnAction = "ClearCellContent"
    End With
    
'    On Error GoTo 0
'/#Context Menu
    '->VBE
'    Call openVBE
End Sub

Sub resetMenu()
          Application.CommandBars("cell").Reset
 End Sub
 
''Private Sub Workbook_Open()
''    nbrMinDistanceTableDescription
''
''End Sub
'Private Sub Workbook_Open()
'
''#Context Menu
'' Menu accessible quand on fait le clique avec le bouton de droit dans une cellule
'Dim cBut As CommandBarButton
'
''    On Error Resume Next
'    'On efface le menu
'    Application.CommandBars("Cell").Reset
'    Set cBut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
'
'    'On le remet
'    With cBut
'       .BeginGroup = False
'       .Caption = "Functions..."
'       .Style = msoButtonIconAndCaption
'       .FaceId = 385
'       .OnAction = "launch_functionDialog"
'    End With
'
'    '#Improvment
'       Set cBut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
'    'On le remet
'    With cBut
'       .BeginGroup = False
'       .Caption = "Euclidian Norme L²"
'       .Style = msoButtonIconAndCaption
'       .FaceId = 521
'       .OnAction = "launch_euclidianNorm2D"
'
'    End With
'        'clear cells
'    Set cBut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
'    With cBut
'        .BeginGroup = False
'        .Style = msoButtonIconAndCaption
'        .FaceId = 536   '1088
'        .Caption = "Delete Content"
'        .OnAction = "ClearCellContent"
'    End With
''    On Error GoTo 0
''/#Context Menu
'    '->VBE
''    Call openVBE
'End Sub
'
''#VBE
'    Private Sub openVBE()
'        Application.VBE.MainWindow.Visible = True
'    End Sub
'
'
''/#VBE

