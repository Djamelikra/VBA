Attribute VB_Name = "mImportCsvFile"
Option Explicit
Sub importCsvFile()
    '
    Dim dlgBox As FileDialog
    Dim importedFile$
    Dim nbrRow&
    Dim lineFromFile$
    Dim lineItem As Variant 'array of $
    Dim i& 'counter
    Dim rangeDestination$ ' in the goal to be adapted !
    Dim nbrColumDestination% ' in the goal to be adapted !
    
    Set dlgBox = Application.FileDialog(msoFileDialogFilePicker)
    
    With dlgBox
        With .Filters
            .Add "CSV", "*.csv", 1
            .Add "Txt", "*.txt", 2
        End With
        .AllowMultiSelect = False
        If .Show Then
            importedFile = .SelectedItems(1)
        End If
    End With
    
    If importedFile <> "" Then
     'open for reading sequence access
        Open importedFile For Input As #1
        nbrRow = 1
        Do Until EOF(1) 'false until end of file is reach then it will turn true
            Line Input #1, lineFromFile 'reads a single line from an open sequential file
            lineItem = Split(lineFromFile, ",") ' or ";" ==> to be adapted !
                For i = 0 To nbrColumDestination
                    Range(rangeDestination).Cells(nbrRow, i + 1) = lineItem(i)
                Next i
            nbrRow = nbrRow + 1
        Loop
        
         'For writing
        Close #1
    End If
    
End Sub
