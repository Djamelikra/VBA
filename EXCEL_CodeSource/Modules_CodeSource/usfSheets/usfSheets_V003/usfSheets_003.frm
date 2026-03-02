VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfSheets 
   Caption         =   "Sheet Choice"
   ClientHeight    =   1815
   ClientLeft      =   645
   ClientTop       =   2325
   ClientWidth     =   5640
   OleObjectBlob   =   "usfSheets_003.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "usfSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOpenVBE_Click()
    'Open VBE
    Application.VBE.MainWindow.Visible = True
    'Close usf
    Me.Hide
    Unload Me
End Sub

Private Sub UserForm_Initialize()
 ' For Each s In ActiveWorkbook.Sheets
 '   Me.ComboBox1.AddItem s.Name
 ' Next s
 Dim i As Integer, n As Integer
 
  Dim temp() As Variant
  For i = 1 To Sheets.Count
    ReDim Preserve temp(1 To i)
    temp(i) = Sheets(i).Name
  Next i
  n = UBound(temp)
  Call Tri(temp, 1, n)
  Me.ComboBox1.List = temp
  Me.ComboBox1.ListIndex = 0
End Sub
Private Sub UserForm_Activate()
'Worksheets in the workbook
    With Me
        .Caption = "There are: " & .ComboBox1.ListCount & " Worksheets in the workbook."
    End With

End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    'to close ?
    Me.Show
    Me.ComboBox1.SetFocus
'    SendKeys "{F4}"  'Drop down
End Sub
Private Sub ComboBox1_Change()
Dim m As String
    m = Me.ComboBox1.Value
  Sheets(m).Select
'Sheets(Me.ComboBox1).Select
End Sub
Sub Tri(a As Variant, gauc As Long, ByVal droi As Long)          ' Quick sort
Dim ref As Variant, g As Long, d As Long, temp As Variant

 ref = a((gauc + droi) \ 2)
 g = gauc: d = droi
 Do
     Do While a(g) < ref: g = g + 1: Loop
     Do While ref < a(d): d = d - 1: Loop
     If g <= d Then
       temp = a(g)
       a(g) = a(d)
       a(d) = temp
       g = g + 1: d = d - 1
     End If
 Loop While g <= d
 If g < droi Then Call Tri(a, g, droi)
 If gauc < d Then Call Tri(a, gauc, d)
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdQuit_Click()
      Application.Quit
End Sub
