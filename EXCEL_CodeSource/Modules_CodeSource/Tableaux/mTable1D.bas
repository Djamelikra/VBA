Attribute VB_Name = "mTable1D"
Option Explicit
'https://learn.microsoft.com/fr-fr/office/vba/api/excel.listobject


Sub makeTable1D()
Dim table As ListObject
Dim rg As Range
Set rg = Cells(2, 1).CurrentRegion


Set table = ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=rg, XlListObjectHasHeaders:=xlYes)

With table
   .Range.HorizontalAlignment = xlCenter  ' Alignement horizontal du contenu des cellules
   .ShowTableStyleRowStripes = False ' Lignes sur couleurs de fond alternées
   .ShowTotals = True  ' Affichage de la ligne de totaux
   .ShowAutoFilterDropDown = False ' Affichage des boutons de filtres automatiques sur les en-têtes
   .TableStyle = "TableStyleLight9" ' Style général (parmi la liste des styles prédéfinis fournis par Excel)
End With
    
End Sub
Sub CreateTableInExcel()

ActiveWorkbook.Sheets("table").ListObjects.Add(xlSrcRange, Range("$A$1:$B$8"), , xlYes).Name = "table1"

End Sub
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
Set Ws = Worksheets("table")
 
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

Sub PopulatingArrayVariable()
'PURPOSE: Dynamically Create Array Variable based on a Given Size

Dim myArray() As Variant
Dim DataRange As Range
Dim cell As Range
Dim x As Long

'Determine the data you want stored
  Set DataRange = ActiveSheet.UsedRange

'Resize Array prior to loading data
  ReDim myArray(DataRange.Cells.Count)

'Loop through each cell in Range and store value in Array
  For Each cell In DataRange.Cells
    myArray(x) = cell.Value
    x = x + 1
  Next cell

'Print values to Immediate Window (Ctrl + G to view)
  For x = LBound(myArray) To UBound(myArray)
    Debug.Print myArray(x)
  Next x

End Sub
