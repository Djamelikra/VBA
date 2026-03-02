Attribute VB_Name = "mTriangleTable"
Option Explicit

Sub lowerTriangleTable()
    Dim i&, j&
    ActiveSheet.Range(Cells(1, 1), Cells(100, 100)).Clear
    
    For i = 1 To 9 'row
        For j = 1 To 9 'col
            If i >= j Then
                Cells(i, j) = i * j
            End If
        Next j
    Next i
    
    MsgBox "Performed ", vbInformation, gAppName
    
End Sub
Sub upperTriangleMultiplicationTable()
    Dim dataTemp$
    Dim dataBuffer$
    Dim i&, j&
    Dim nbrDigit As Byte
    
    'This const can be adapted:
    Const NBR_FINAL As Byte = 10
    
'    ActiveSheet.Range(Cells(1, 1), Cells(100, 100)).Clear
    
    Select Case NBR_FINAL
        Case Is < 10: nbrDigit = 3
        Case 10 To 31: nbrDigit = 4
        Case 31 To 100: nbrDigit = 5
        Case Else: MsgBox "Number too large !", vbExclamation, gAppName
                Exit Sub
    End Select
    
    dataBuffer = String$(nbrDigit, " ")
    
    For i = 1 To NBR_FINAL
        dataTemp = Right(dataBuffer & i, nbrDigit)
            For j = 2 To NBR_FINAL
                If j < i Then
                    dataTemp = dataTemp & dataBuffer
                Else
                    dataTemp = dataTemp & Right(dataBuffer & j * i, nbrDigit)
                End If
            Next j
            Debug.Print dataTemp
    Next i
    
    
    
    
'    MsgBox "Performed ", vbInformation, gAppName
    
End Sub
