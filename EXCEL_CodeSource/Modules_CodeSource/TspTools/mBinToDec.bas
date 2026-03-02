Attribute VB_Name = "mBinToDec"
Option Explicit
Public Function BinToDec(Binary As String) As Double
Dim n As Double
Dim s As Double

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ (s - 1)))
    Next s

    BinToDec = n
End Function
Function BinaryToDecimal(sBinary As String, sSign As String) As Double
Attribute BinaryToDecimal.VB_Description = "Binary To Decimal Conversion; Maximum number of gigits: 31 Only!"
'
     
    Dim i As Double
    Dim lReturn As Double
    Dim lBit As Double
   
    Const sNEGATIVE As String = "1"
    Const sOFF As String = "0"
   
    For i = 1 To Len(sBinary)
        If sSign = sNEGATIVE Then
            If Mid$(sBinary, i, 1) = sOFF Then
                lBit = 1
            Else
                lBit = 0
            End If
        Else
            lBit = Val(Mid$(sBinary, i, 1))
        End If

       lReturn = lReturn + (lBit * (2 ^ (Len(sBinary) - i)))
       
        
    Next i
   
    If sSign = sNEGATIVE Then
        BinaryToDecimal = -(lReturn + 1)
    Else
        BinaryToDecimal = lReturn
    End If
   
End Function
Function BinaryToDecimalNum(ByVal sBinary As Double, sSign As Byte) As Double
'To do : !!!!!!!!! to fix !!!!!!!! not performs !!!!!!
     
    Dim i As Double
    Dim lReturn As Double
    Dim lBit As Double
   
    Const sNEGATIVE As String = "1"
    Const sOFF As String = "0"
   
    For i = 1 To Len(CStr(sBinary))
        If sSign = sNEGATIVE Then
            If Mid$(sBinary, i, 1) = sOFF Then
                lBit = 1
            Else
                lBit = 0
            End If
        Else
            lBit = Val(Mid$(sBinary, i, 1))
        End If

       lReturn = lReturn + (lBit * (2 ^ (Len(sBinary) - i)))
       
        
    Next i
   
    If sSign = sNEGATIVE Then
        BinaryToDecimalNum = -(lReturn + 1)
        
        
    Else
        BinaryToDecimalNum = lReturn
        
    End If
   
End Function
Sub test_BinToDec()
 
    Dim strBin$
    
    strBin = "10101010101010101010101010101010101010101010101010101010101010101010101010101"
    MsgBox BinToDec(strBin), vbInformation, "Number of digits = " & Len(strBin)
    
End Sub

Sub test_BinaryToDecimal()
    Const sNEGATIVE As String = "1"
    Const sOFF As String = "0"
    Dim strBin$
    
    strBin = "101010101010101010101010101010101010101010101010101010101010101010101010101"
    MsgBox BinaryToDecimal(strBin, sOFF), vbInformation, "Number of digits = " & Len(strBin)
    
End Sub
Sub test_BinaryToDecimalNum()
    Const sNEGATIVE As String = "1"
    Const sOFF As String = "0"
    Dim strBin$
    
    strBin = 10
    MsgBox BinaryToDecimalNum(strBin, sOFF), vbInformation, "Number of digits = " & Len(strBin)
    
End Sub

