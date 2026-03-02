VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfImportExportCsv 
   Caption         =   "Data Import / Export"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14190
   OleObjectBlob   =   "usfImportExportCsv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfImportExportCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1 048 576 lignes et 16 384 colonnes
Dim startRow& 'ligne début
Dim startColumn% 'colonne début
Dim endRow& 'ligne fin
Dim endColumn% 'colonne fin
Dim currentRow& 'ligne en cours
Dim currentColumn% 'colonne en cours

Private Sub cmdImport_Click()
'+csv
    Dim selectedFile$ 'fichier choisi
    selectedFile = Application.GetOpenFilename("Text Files (*.txt), *.txt, Fichiers CSV (*.csv), *.csv", 2, "Select a CSV file")
    If LCase$(selectedFile <> "false" And selectedFile <> "0") Then
        lstFiles.AddItem (selectedFile)
    End If
    
    
End Sub
Private Sub cmdExport_Click()
'
    Dim exportFile$
    Dim i As Byte
    
    startRow = 2
    startColumn = 2
    currentRow = startRow
    currentColumn = startColumn
    
    Cells.Clear
    
    For i = 0 To lstFiles.ListCount - 1
        Call readFile(lstFiles.List(i))
    Next i
    
    Call dataProcessing
    
    '+csv
    exportFile = Application.GetSaveAsFilename(, "Text Files (*.txt), *.txt, Fichiers CSV (*.csv), *.csv", 2, "Save the export file", "Storage CSV")
    
    txtFilePath.Value = exportFile
    
    Call writeFile(exportFile)
    
    
    

End Sub

Private Sub readFile(fullNameFile$)
'    MsgBox fullNameFile, vbInformation, gAppName & " " & gCR
    Dim start%
    Dim position%
    Dim textFile$
    Dim buffer$
    
    'open for reading sequence access
    Open fullNameFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, textFile
            start = 1
            position = 1
            Do While (position <> 0)
                position = InStr(start, textFile, ";", vbTextCompare)
                If position = 0 Then
                    buffer = Mid$(textFile, start)
'                   Sheets("Import").Cells(currentRow, currentColumn).Value = buffer
                   shDataImport.Cells(currentRow, currentColumn).Value = buffer
                   
                    Exit Do
                Else
                    buffer = Mid(textFile, start, position - start)
                End If
                
                Sheets("Import").Cells(currentRow, currentColumn).Value = buffer
                start = position + 1
                currentColumn = currentColumn + 1
            Loop
            
            currentColumn = startColumn
            currentRow = currentRow + 1
            
        Loop
    Close #1
    
End Sub

Private Sub writeFile(fileName$)
'Export From excel to csv

Dim dataRow& '
Dim dataColumn% '
Dim textBuffer$

dataRow = startRow
dataColumn = startColumn


If LCase(fileName) <> False Then
    'For writing
    Open fileName For Output As #1
    
           
        While Cells(dataRow, dataColumn).Value <> ""
            
            While Cells(dataRow, dataColumn).Value <> ""
                textBuffer = textBuffer & Cells(dataRow, dataColumn).Value & ";"
                dataColumn = dataColumn + 1
            Wend
            
            Print #1, textBuffer
            textBuffer = ""
            dataColumn = startColumn
            dataRow = dataRow + 1
        
        Wend
    
    Close #1

End If
    

End Sub
Private Sub dataProcessing()
'processing : removal of duplicates
    Dim dataRow& '
    Dim dataColumn% '
    
    dataRow = startRow
    dataColumn = startColumn
    
    Cells(dataRow, dataColumn).Sort Cells(dataRow, dataColumn), xlAscending, Header:=xlNo
    
    While Cells(dataRow, dataColumn).Value <> ""
    
         If (Cells(dataRow, dataColumn).Value = Cells(dataRow - 1, dataColumn).Value) Then
             Cells(dataRow, dataColumn).EntireRow.Delete
             dataRow = dataRow - 1
         End If
         
        dataRow = dataRow + 1
    
    Wend
    
    
End Sub
Private Sub cmdCloseForm_Click()
'
    lstFiles.Clear
    Unload Me
    
End Sub

