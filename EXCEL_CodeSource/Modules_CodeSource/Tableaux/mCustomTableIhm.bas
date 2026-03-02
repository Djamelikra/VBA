Attribute VB_Name = "mCustomTableIhm"
Option Explicit
'mCustomTableIhm


Sub makeCustomTable()
Dim oWsh As Worksheet
Dim tableName$
'Dim wshName$ 'Worksheet's name
Dim workRngNE As Range 'start range
Dim workRngSW As Range 'end range
'xlSrcExternal 0 Source de données externes (site Microsoft Windows SharePoint Services).
'xlSrcQuery 3 Requête
'xlSrcRange 1 Plage
'xlSrcXml 2 XML
 
    tableName = Application.InputBox("Enter the table name", "Table Name", "customTable")
'Set oWsh = Worksheets("table3")
    Set oWsh = Worksheets(tableName)
    
 
'Un exemple qui prend en compte les cellules voisines à A1
'With oWsh
'    .ListObjects.Add(xlSrcRange, .Range("$A$1").CurrentRegion, , xlYes).Name = tableName
'    .ListObjects(tableName).TableStyle = "TableStyleMedium5"
'End With
 
 'todo .../..
'
' Set WorkRng = Application.Selection
'Set WorkRng = Application.InputBox("Range", , WorkRng.Address, Type:=8)
  'todo .../..
 Set workRngNE = Application.Selection
 Set workRngNE = Application.InputBox("Range", , workRngNE.Address, Type:=8)
   MsgBox workRngNE.Address
    Set workRngSW = Application.Selection
 Set workRngSW = Application.InputBox("Range", , workRngSW.Address, Type:=8)
   MsgBox workRngSW.Address
'Un autre exemple qui prend en compte une plage spécifique.
  
    

With oWsh
'    .ListObjects.Add(xlSrcRange, .Range("$B$5:$F$30"), , xlYes).Name = tableName
' .ListObjects.Add(xlSrcRange, .Range(workRng), , xlYes).Name = tableName
' .ListObjects.Add(xlSrcRange, .Range(workRng.Address).CurrentRegion, , xlYes).Name = tableName
'.ListObjects.Add(xlSrcRange, .Range(Cells(1, 1), Cells(10, 10)), xlYes).Name = tableName
 .ListObjects.Add(xlSrcRange, .Range((workRngNE.Address), (workRngSW.Address)), , xlYes).Name = tableName
    .ListObjects(tableName).TableStyle = "TableStyleMedium5"
'    .ListObjects(tableName).AutoFilter.ShowAllData
End With

     Range("customTable[[#Headers],[Column1]]").Select
    Selection.AutoFilter
    
     

     MsgBox "Done !", vbInformation, gAppName
     
     oWsh.Activate
     
     
End Sub


