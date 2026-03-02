Attribute VB_Name = "mTable"
Option Explicit
Sub sbCreatTable()

    'Create Table in Excel VBA
    Sheet1.ListObjects.Add(xlSrcRange, Range("A1:D10"), , xlYes).Name = "myTable1"
    Range("B2").Select
    Selection.AutoFilter

End Sub
Sub CreationTableau()
Dim Ws As Worksheet
Dim NomTable As String
 
'xlSrcExternal 0 Source de données externes (site Microsoft Windows SharePoint Services).
'xlSrcQuery 3 Requête
'xlSrcRange 1 Plage
'xlSrcXml 2 XML
 
NomTable = "TableEx1"
Set Ws = Worksheets("table")
 
'Un exemple qui prend en compte les cellules voisines à A1
With Ws
    .ListObjects.Add(xlSrcRange, .Range("$A$1").CurrentRegion, , xlYes).Name = NomTable
    .ListObjects(NomTable).TableStyle = "TableStyleMedium5"
End With
 
 
'Un autre exemple qui prend en compte une plage spécifique.
'With Ws
    '.ListObjects.Add(xlSrcRange, .Range("$B$5:$F$30"), , xlYes).Name = NomTable
    '.ListObjects(NomTable).TableStyle = "TableStyleMedium5"
'End With

    MsgBox "Done !", vbInformation, gAppName
    
End Sub

Sub CreationTableauSpecifique()
Dim Ws As Worksheet
Dim NomTable As String
 
'xlSrcExternal 0 Source de données externes (site Microsoft Windows SharePoint Services).
'xlSrcQuery 3 Requête
'xlSrcRange 1 Plage
'xlSrcXml 2 XML
 
NomTable = "TableEx2"
Set Ws = Worksheets("table3")
 
'Un exemple qui prend en compte les cellules voisines à A1
'With Ws
'    .ListObjects.Add(xlSrcRange, .Range("$A$1").CurrentRegion, , xlYes).Name = NomTable
'    .ListObjects(NomTable).TableStyle = "TableStyleMedium5"
'End With
 
 
'Un autre exemple qui prend en compte une plage spécifique.
With Ws
'    .ListObjects.Add(xlSrcRange, .Range("$B$5:$F$30"), , xlYes).Name = NomTable
.ListObjects.Add(xlSrcRange, .Range(Cells(1, 1), Cells(10, 10)), xlYes).Name = NomTable
    .ListObjects(NomTable).TableStyle = "TableStyleMedium5"
End With

     MsgBox "Done !", vbInformation, gAppName
     
End Sub

