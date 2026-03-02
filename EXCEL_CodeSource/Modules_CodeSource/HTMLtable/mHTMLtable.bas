Attribute VB_Name = "mHTMLtable"
Option Explicit
'mHTMLtable

Public Function makeHTMLTable(rInput As Range, Optional bHeaders As Boolean = True) As String
    
    Dim rRow As Range
    Dim rCell As Range
    Dim sReturn As String
    
    sReturn = " < Table > """
    
    If bHeaders Then
        sReturn = sReturn & "<tr><td> </td>"
        
        For Each rCell In rInput.Rows(1).Cells
            sReturn = sReturn & " < td > " & Chr$(rCell.Column + 64) & "</td>"
        Next rCell
        
        sReturn = sReturn & "</tr>"
    End If
    
    For Each rRow In rInput.Rows
        sReturn = sReturn & " < tr > """
        
        If bHeaders Then
            sReturn = sReturn & " < td > " & rRow.Row & "</td>"
        End If
        
        For Each rCell In rRow.Cells
            sReturn = sReturn & " < td > " & rCell.Text & "</td>"
        Next rCell
        
        sReturn = sReturn & "</tr>" & vbNewLine
    Next rRow
    
    sReturn = sReturn & "</table>"
    
    makeHTMLTable = sReturn
    
End Function
