VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Private Sub Workbook_Open()
'    usfImportExportCsv.Show
'End Sub
Private Sub Workbook_Open()
'resetmenu
'appel la macro qui effacera la commande du menu contextuel
'crée lors de la session précédente

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
    
    'clear cells
    Set Cbut = Cpop1.Controls.Add(Type:=msoControlButton)
    With Cbut
        .FaceId = 1088
        .Caption = "Delete"
        .OnAction = "ClearCellContent"
    End With
End Sub
Sub resetmenu()
          Application.CommandBars("cell").Reset
 End Sub



