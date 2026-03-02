Attribute VB_Name = "mContextMenu"
Option Explicit
'mContextMenu <=> ThisWorkbook
Dim Ch As Range
Sub majuscule()
  For Each Ch In Selection
  'Pour chaque cellule Ch de ma sélection Selection
  'Selection permet de renvoyer l'objet sélectionné dans la fenêtre active
    If Not Ch.HasFormula Then
    'la propriété HasFormula de l'objet range permet de tester qu'aucune
    ' cellules de la plage ne contiennent de formule, dans ce cas la conversion est ignorée
      Ch.Value = UCase(Ch.Value) 'Upper Case = MAJUSCULE
    End If
  Next Ch
End Sub
'*************************************************************
Sub minuscule()
  For Each Ch In Selection
    If Not Ch.HasFormula Then
      Ch.Value = LCase(Ch.Value) 'Lower Case = minuscule
    End If
  Next Ch
End Sub

'*************************************************************
 Sub nompropre()
  For Each Ch In Selection
    If Not Ch.HasFormula Then
      Ch.Value = Application.Proper(Ch.Value)
      'si une fonction Excel n'a pas d'équivalent en VBA
      'utiliser le la fonction Excel (anglais) comme une méthode de l'objet application
    End If
  Next Ch
End Sub
'***************************************************************
Sub ClearCellContent()
'1088; 536
'Clear content
    Selection.ClearContents
    ActiveCell.Select
End Sub
