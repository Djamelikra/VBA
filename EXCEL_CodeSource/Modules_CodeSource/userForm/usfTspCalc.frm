VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfTspCalc 
   Caption         =   "TSP calc"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10110
   OleObjectBlob   =   "usfTspCalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfTspCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim retNbrNodes%
Private Sub UserForm_Initialize()
'CONSTRUCTOR
    Dim i As Byte
    'params
    cmdCalc.SetFocus
    With cboNbrCities
        .Style = fmStyleDropDownList
        For i = 3 To 23
            .AddItem (i)
        Next i
        
        .ListIndex = 0 'first item
        
    End With
End Sub
Private Sub cboNbrCities_Change()
      cmdCalc.SetFocus
End Sub

Private Sub cmdCalc_Click()
    'Main program
    retNbrNodes = CInt(cboNbrCities.Value)
    ''possible cycles
    If cboNbrCities.ListIndex <> -1 Then
        txtNbrPaths.Text = recursiveFactorial(retNbrNodes)
        ''possible cycles
        txtNbrEdges = nbrEdges(retNbrNodes)
        ''different cycles
        txtNbrDiffPaths = nbrDifferentPaths(retNbrNodes)
        ''candidate cycles
        txtNbrCandidPaths = nbrCandidatPaths(retNbrNodes)
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function nbrEdges(nbrNodes) As Long
    nbrEdges = nbrNodes * (nbrNodes - 1) / 2
End Function
Private Function recursiveFactorial(ByVal n As Double) As Double

     If n > 0 Then recursiveFactorial = n * factorialRecursion(n - 1)
     If n = 0 Then recursiveFactorial = 1

 End Function
Private Function nbrDifferentPaths(ByVal nbrVertex As Long) As Double
'Double because the number cases can be huge !
    nbrDifferentPaths = factoRecurs(nbrVertex - 1)
End Function

Private Function nbrCandidatPaths(ByVal nbrVertex As Long) As Double
'Double because the number cases can be huge !
    nbrCandidatPaths = factoRecurs(nbrVertex - 1) / 2
End Function
Private Function DiracNumber(nbrVertex As Long) As Long

    DiracNumber = (nbrVertex) \ 2

End Function








