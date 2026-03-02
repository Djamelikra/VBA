Attribute VB_Name = "mShape"
Option Explicit
'mShape

Sub AddShape()
    Dim oWshapes As Worksheet
'    Set oWshapes = Worksheets("shapes")
    
    Set oWshapes = ActiveSheet
    
    
'    oWshapes.Select
    
    
    With oWshapes.Shapes.AddShape(msoShapeRectangle, 40, 80, 140, 50)
        .Name = "rectShape"
        .TextFrame.Characters.Text = "Max rho"
            With .TextFrame
                .AutoSize = True
            
            End With
                With .TextFrame.Characters
'                    .Font.FontStyle = "Elephant"
                    .Font.Name = "Algerian"
                    .Font.Size = 24
                    .Font.Color = vbGreen
                End With
    End With
End Sub

Sub AddShapeCustom()
    Dim oWshapes As Worksheet
'    Set oWshapes = Worksheets("shapes")
    Dim oShapRectang As Shape
    Dim sPathImage As String
    
    Dim L As Integer, T As Integer, H As Integer, W As Integer
    Dim xy As String
    xy = "K115"
    
    'Dimensions et position de la zone de texte
    H = Range(xy).Height * 4 '<-- hauteur
    W = Range(xy).Left - Range(xy).Left '<-- largeur
    L = Range(xy).Left '<-- position horizontale
    T = Range(xy).Top '<-- position verticale'Insertion de la zone de texte

    sPathImage = ThisWorkbook.Path & "/images/IndexRight.ico"
    
    
    Set oWshapes = ActiveSheet
'    ActiveCell.Select
'    Range("B29").Select
    
'    oWshapes.Select
'    Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, 100, 100, 300, 200)
     Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    
    With oShapRectang
        With .Fill
            .ForeColor.RGB = RGB(0, 255, 0)
'            .BackColor.RGB = RGB(255, 0, 0)
'            .UserPicture sPathImage
        End With

        .Name = "rectangShape"
        .TextFrame.Characters.Text = "Max rho"
            With .Line
                .ForeColor.RGB = RGB(255, 0, 0)
                .Weight = 15
            End With
            With .TextFrame
                .AutoSize = True
            
            End With
                With .TextFrame.Characters
                    .Font.Bold = True
'                    .Font.Name = "Algerian"
                    .Font.Size = 36
                    .Font.Color = vbYellow
                End With
        
    End With
End Sub

Sub AddRectangle()

   Dim oWshapes As Worksheet
    Dim oShapRectang As Shape
    Set oWshapes = ActiveSheet
    
    
 With oWshapes.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
'            .ForeColor.RGB = RGB(128, 0, 0)
            .BackColor.RGB = RGB(170, 170, 170)
'            .TwoColorGradient msoGradientHorizontal, 1
End With
End Sub
Sub ToutesFormes()
Dim sh As Shape
Dim i As Integer, lf As Integer, tp As Integer
' Sélectionnez la feuille de calcul
ThisWorkbook.Worksheets("shapes").Activate
' Pas de grille
ActiveWindow.DisplayGridlines = False
' À vider
For i = ActiveSheet.Shapes().Count To 1 Step -1
ActiveSheet.Shapes(i).Delete
Next i
' Valeurs de départ
lf = 5
tp = 5
' Toutes les formes différentes
For i = 1 To 137
Set sh = ActiveSheet.Shapes(). _
AddShape(i, lf, tp, 30, 30)
' Mise en page
sh.Line.Weight = 1
sh.Line.ForeColor.RGB = RGB(0, 0, 0)
sh.Fill.ForeColor.RGB = RGB(255, 255, 255)
With sh.TextFrame.Characters
.Font.Color = vbBlack

.Font.Size = 7
.Text = i
End With
' Position suivante
lf = lf + 35
If i Mod 15 = 0 Then
lf = 5
tp = tp + 35
End If
Next i
End Sub
Sub DeleteAllShapes()
    Do Until ActiveSheet.Shapes.Count = 0
        ActiveSheet.Shapes(1).Delete
    Loop
End Sub
Sub openWorkspace()
    Dim workPath As String
    workPath = ThisWorkbook.Path
    MsgBox workPath, vbInformation, "The Workspace Directory: "
    Shell "C:\windows\explorer.exe " & workPath, vbMaximizedFocus
End Sub

Sub paths()
    Dim strPathImg As String
    Dim strPath$
    strPathImg = ThisWorkbook.Path & "/images/IndexRight.ico"
    
    MsgBox strPathImg


    
    
End Sub
