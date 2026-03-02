Attribute VB_Name = "mVBE"
Option Explicit
Option Compare Text

Sub testAjout()
'"laProcedure" est la macro dont vous souhaitez numéroter les lignes
NumerotationLignesProcedure "Module1", "laProcedure"
End Sub


Sub NumerotationLignesProcedure(nomModule As String, nomMacro As String)
'
'Cet exemple ne gère pas les erreurs s'il existe déja une numérotation
'
Dim Debut As Integer, Lignes As Integer, x As Integer
Dim Texte As String, strVar As String

With ThisWorkbook.VBProject.VBComponents(nomModule).CodeModule
    Debut = .ProcStartLine(nomMacro, 0)
    Lignes = .ProcCountLines(nomMacro, 0)
End With

For x = Debut + 2 To Debut + Lignes - 1
    With ThisWorkbook.VBProject.VBComponents(nomModule).CodeModule
        Texte = .Lines(x, 1)
        strVar = Application.WorksheetFunction.Substitute(Texte, " ", "")
        strVar = Application.WorksheetFunction.Substitute(strVar, vbCrLf, "")
        strVar = Application.WorksheetFunction.Substitute(strVar, vbTab, "")
        
        'Adaptez les filtres en fonction de vos projets.
        'Remarque: les arguments PrivateFunction et PublicFunction sont volontairement accolés.
        '
        If strVar <> "" And _
        Left(strVar, 3) <> "Sub" And _
        Left(strVar, 10) <> "PrivateSub" And _
        Left(strVar, 9) <> "PublicSub" And _
        Left(strVar, 8) <> "Function" And _
        Left(strVar, 15) <> "PrivateFunction" And _
        Left(strVar, 14) <> "PublicFunction" And _
        Right(ThisWorkbook.VBProject.VBComponents(nomModule). _
                CodeModule.Lines(x - 1, 1), 1) <> "_" _
        Then .ReplaceLine x, x & " " & Texte
    End With
Next
End Sub
Sub testSuppression()
supprimeNumerotationLignes "Module1", "laProcedure"
End Sub


Sub supprimeNumerotationLignes(nomModule As String, nomMacro As String)
Dim Debut As Integer, Lignes As Integer, x As Integer
Dim Texte As String, strVar As String
Dim Valeur As Integer

With ThisWorkbook.VBProject.VBComponents(nomModule).CodeModule
    Debut = .ProcStartLine(nomMacro, 0)
    Lignes = .ProcCountLines(nomMacro, 0)
End With

For x = Debut + 2 To Debut + Lignes - 1
    With ThisWorkbook.VBProject.VBComponents(nomModule).CodeModule
        Texte = .Lines(x, 1)
        
            Valeur = Val(Texte)
            If Valeur <> 0 Then
                strVar = Mid(Texte, Len(CStr(Valeur)) + 2)
                Else
                strVar = Texte
            End If
            
            .ReplaceLine x, strVar
    End With
Next
End Sub
