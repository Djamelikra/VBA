Attribute VB_Name = "mBruteForce"
Option Explicit

Sub Makro1()
'Todo !

Dim a1, a2, a3, a4, a5, a6, a7, a8, a9 As Integer
Dim n As Long

Range("a1").Select

n = 1

For a1 = 1 To 9
 For a2 = 1 To 9
 If (a2 <> a1) Then
  For a3 = 1 To 9
  If ((a3 <> a1) And (a3 <> a2)) Then
   For a4 = 1 To 9
    If ((a4 <> a1) And (a4 <> a2) And (a4 <> a3)) Then
    For a5 = 1 To 9
    If ((a5 <> a1) And (a5 <> a2) And (a5 <> a3) And (a5 <> a4)) Then
     For a6 = 1 To 9
     If ((a6 <> a1) And (a6 <> a2) And (a6 <> a3) And (a6 <> a4) And (a6 <> a5)) Then
      For a7 = 1 To 9
      If ((a7 <> a1) And (a7 <> a2) And (a7 <> a3) And (a7 <> a4) And (a7 <> a5) And (a7 <> a6)) Then
       For a8 = 1 To 9
       If ((a8 <> a1) And (a8 <> a2) And (a8 <> a3) And (a8 <> a4) And (a8 <> a5) And (a8 <> a6) And (a8 <> a7)) Then
        For a9 = 1 To 9
         If ((a9 > a1) And (a9 <> a2) And (a9 <> a3) And (a9 <> a4) And (a9 <> a5) And (a9 <> a6) And (a9 <> a7) And (a9 <> a8)) Then
         Cells(n, 1).Value = a1
         Cells(n, 2).Value = a2
         Cells(n, 3).Value = a3
         Cells(n, 4).Value = a4
         Cells(n, 5).Value = a5
         Cells(n, 6).Value = a6
         Cells(n, 7).Value = a7
         Cells(n, 8).Value = a8
         Cells(n, 9).Value = a9
         n = n + 1
         End If
        Next a9
        End If
       Next a8
      End If
      Next a7
     End If
     Next a6
    End If
    Next a5
   End If
   Next a4
  End If
  Next a3
 End If
 Next a2
Next a1
        
End Sub


