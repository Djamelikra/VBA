VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdRun_Click()
    Dim intResult As Long
    'displays the color dialog
    Application.Dialogs(xlDialogEditColor).Show 40, 100, 100, 200
    
    'gets the color selected by the user
    intResult = ThisWorkbook.Colors(40)
    'changes the fill color of cell A1
    Range("A1").Interior.Color = intResult
End Sub
