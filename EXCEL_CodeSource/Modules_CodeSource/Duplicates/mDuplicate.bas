Attribute VB_Name = "mDuplicate"
Option Explicit
'mDuplicate
'###1
Sub identifyDuplicate()
    Dim dataRange As Range
    Dim Cell As Range
    Dim Un As Collection
    Dim i As Integer
    
    Set dataRange = Application.Selection
    Set dataRange = Application.InputBox("Range", "dataRange", dataRange.Address, Type:=8)
    Set Un = New Collection
  
    On Error Resume Next
    
    'Boucle sur la plage de cellule
      i = 0
    For Each Cell In dataRange
        'Pour ne pas prendre en compte les cellules vides
        If Cell <> "" Then
            'Ajoute le contenu de la cellule dans la collection
            Un.Add Cell, CStr(Cell)
            
            'Si la procédure renvoie une erreur, cela signifie que l'élément
            'existe déjà dans la collection et donc qu'il s'agit d'un doublon.
            'Dans ce cas la macro colorie la cellule en vert.
          
            If Err <> 0 Then
                Cell.Interior.ColorIndex = 4
                i = i + 1
            End If
            'Efface toutes les valeurs de l'objet Err.
            Err.Clear
        End If
        
        
    Next Cell
    MsgBox i, vbInformation
    If i > 0 Then
        MsgBox "Number of duplicates: " & i, vbExclamation
        Else
        MsgBox "No duplicates ", vbInformation
    End If
    
    Set Un = Nothing
End Sub


'###2
Sub doublons_et_lignes_vides()
 Dim choix As String
 Dim choix2 As String
 Dim Test As Single
 Dim der_ligne As Long
 Dim Ligne As Long
 Dim nb As Long
 Dim Compteur As Long
 Dim Contenu As String
 Dim i As Long
 Dim res_test As String
 
    
    choix = InputBox("Avant d'utiliser cet outil, n'oubliez pas d'enregistrer votre fichier !" & Chr(10) & Chr(10) & "Choisissez l'action qui vous intéresse :" & Chr(10) & Chr(10) & "1. Colorer les doublons (colorer la cellule)" & Chr(10) & "2. Colorer les doublons (colorer la ligne entière)" & Chr(10) & "3. Effacer les doublons (en laissant la ligne vide)" & Chr(10) & "4. Supprimer les doublons (ligne entière)" & Chr(10) & "5. Supprimer les lignes vides" & Chr(10) & Chr(10) & "Entrez le n° de l'action et cliquez sur OK :", "Gestion des doublons - Excel-Pratique.com")
    If choix = "" Then Exit Sub
 
    choix2 = ""
    If choix = 1 Or choix = 2 Or choix = 3 Or choix = 4 Then choix2 = InputBox("Entrez la lettre de la colonne où les doublons doivent être recherchés :", "Gestion des doublons - Excel-Pratique.com")
    If choix = 5 Then choix2 = InputBox("Entrez la lettre de la colonne à prendre en compte (si la cellule de cette colonne est vide, la ligne sera supprimée) :", "Gestion des doublons - Excel-Pratique.com")
    If choix2 = "" Then Exit Sub
 
    Application.ScreenUpdating = False
    Test = Timer
 
    der_ligne = Range(choix2 & Rows.Count).End(xlUp).Row
 
    Dim tab_cells()
    ReDim tab_cells(der_ligne - 1)
 
    For Ligne = 1 To der_ligne
        tab_cells(Ligne - 1) = Range(choix2 & Ligne)
    Next
 
    nb = 0
    If choix = 4 Or choix = 5 Then Compteur = 0
 
    For Ligne = 1 To der_ligne
        Contenu = tab_cells(Ligne - 1)
 
        If (choix = 1 Or choix = 2) And Contenu <> "" Then 'Colorer doublons
            For i = 1 To der_ligne
                If Contenu = tab_cells(i - 1) And Ligne <> i Then 'Si doublon
                    nb = nb + 1
                    If choix = 1 Then
                        Range(choix2 & Ligne).Interior.ColorIndex = 3
                    Else
                        Range(Ligne & ":" & Ligne).Interior.ColorIndex = 3
                    End If
                    Exit For
                End If
            Next
        End If
 
        If (choix = 3 Or choix = 4) And Ligne > 1 And Contenu <> "" Then 'Effacer/supprimer doublons
            For i = 1 To Ligne - 1
                If Contenu = tab_cells(i - 1) Then 'Si doublon
                    nb = nb + 1
                    If choix = 3 Then
                        Range(Ligne & ":" & Ligne).ClearContents
                    Else
                        Range(Ligne + Compteur & ":" & Ligne + Compteur).Delete
                        Compteur = Compteur - 1
                    End If
                    Exit For
                End If
            Next
        End If
 
        If choix = 5 And Contenu = "" Then 'Lignes vides
            Range(Ligne + Compteur & ":" & Ligne + Compteur).Delete
            Compteur = Compteur - 1
            nb = nb + 1
        End If
    Next
 
    res_test = Format(Timer - Test, "0" & Application.DecimalSeparator & "000")
    Application.ScreenUpdating = True
 
    If nb = 0 And choix = 5 Then
        MsgBox "Aucune ligne vide trouvée ...", 64, "Résultat"
    ElseIf nb = 0 Then
        MsgBox "Aucun doublon trouvé dans la colonnne " & UCase(choix2) & " ...", 64, "Résultat"
    ElseIf choix = 5 Then
        MsgBox nb & " lignes supprimées (en " & res_test & " secondes)", 64, "Résultat"
    ElseIf choix = 4 Then
        MsgBox nb & " doublons supprimés (en " & res_test & " secondes)", 64, "Résultat"
    ElseIf choix = 3 Then
        MsgBox nb & " doublons effacés (en " & res_test & " secondes)", 64, "Résultat"
    Else
        MsgBox nb & " doublons passés en rouge (en " & res_test & " secondes)", 64, "Résultat"
    End If
 
End Sub

