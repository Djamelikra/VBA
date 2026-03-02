VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit 'Sheet11 'TSP_14
Sub FillFullMatrixWithOrdinalData_001()
' algo_1
'Not generic
  Range("B37:O50").Select
    Selection.Copy
    Range("B55").Select
    ActiveSheet.Paste
    Range("A69").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
     Call messagePerf
     
     
End Sub
Sub findAndLighTheValMin1_002()
'algo_2
'Not generic
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    Dim xTitleId$
    Dim valMin%
    Dim customGreen$
    
    customGreen = RGB(59, 255, 148)
    
      
    On Error Resume Next
    xTitleId = "Looking for min value = 1 "
'    Set WorkRng = Application.Selection
Set WorkRng = Range("B77:O90") '=> Is the data proccess Range
    'Range("B66:Q81")
    
'    Set WorkRng = Application.InputBox(DATA_RANGE, xTitleId, WorkRng.Address, Type:=8)
    
'    valMin = Application.WorksheetFunction.Min(WorkRng)
  valMin = 1
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            If Rng.Interior.Color = vbYellow Then
                Rng.Interior.Color.Index = 45
            End If
'                Rng.Value = ""
                Rng.Interior.Color = customGreen
        End If
    Next
    
    
'    Call messagePerf
    
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub FindAndRemoveSupNumber_003()
'Generic
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrRemoved%
      Dim remainingValues%
      
      'there are five values left
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrRemoved = 0
    For Each Rng In WorkRng
       
        If Rng.Value > CInt(trackedValue) Then
            Rng.Value = ""
        nbrRemoved = nbrRemoved + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrRemoved - [A89] 'less the diagonal
    
   
    MsgBox nbrRemoved & " items out of " & totalValues & " values have been removed." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
End Sub
Sub fillWithDataProcessed_004()
' algo_4
''Not generic
Dim SMsg As String
    SMsg = "Eurék@!_fillWithDataProcessed_004"
    
Application.ScreenUpdating = False
    'Data source
    Range("B76:U95").Select
    Selection.Copy
    
    'Data destination
    Range("B106").Select
    ActiveSheet.Paste
    Range("B131").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
    ActiveWorkbook.Save
    
       'Delete Arrow shapes used in the algorithm
    Call deleteUpDownArrowShapes
    
    Beep
'
       Call msgBoxTempMsgWait(SMsg, 3)
    
End Sub
Sub fillWithDataProcessLayer2_009()
' algo_9
''Not generic
Dim SMsg As String
    SMsg = "Eurék@!_fillWithDataProcessLayer2_009"
    
Application.ScreenUpdating = False
    'Data source
    Range("Z76:AS95").Select
    Selection.Copy
    
    'Data destination
    Range("B106").Select
    ActiveSheet.Paste
    Range("B131").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
    
'    ActiveWorkbook.Save
    
    Beep
'
       Call msgBoxTempMsgWait(SMsg, 3)
End Sub
Sub fillWithDataProcessScalar()
' algo_9
''Not generic
Dim SMsg As String
    SMsg = "Eurék@! ->  fillWithDataProcessScalar"
    
Application.ScreenUpdating = False
    'Data source
    Range("Y2:AR21").Select
    Selection.Copy
    
    'Data destination
    Range("B106").Select
    ActiveSheet.Paste
    Range("B131").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
    
'    ActiveWorkbook.Save
    
    Beep
'
       Call msgBoxTempMsgWait(SMsg, 3)
End Sub
Sub findAndLighValMinimumParam_005()
'alg_05
'Generic
  'Color in orange a new value choosed as min
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrUpdated%
      Dim remainingValues%
      
        Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
      Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrUpdated = 0
    For Each Rng In WorkRng
       
        If Rng.Value = CInt(trackedValue) Then
          
        Rng.Interior.Color = orangeColor
        nbrUpdated = nbrUpdated + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrUpdated
   
    MsgBox nbrUpdated & " items out of " & totalValues & " values have been updated." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
    
End Sub

Sub findAndLighValMinSecondSG_0061()
'alg_06_1
'Generic
  '2024/03/05
    'To do : insert a flag with a shape please !
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
      Dim msgMin$
      Dim msgMaxSingle$
      Dim msgMaxMultiple$
      Dim msgRhoCalc$
      Dim msgNoChiral$
      Dim iCount As Byte
      
     msgNoChiral = "IF The node is No Chiral !"
     msgMin = "Calculate the next Min(x,y) AND Take Rho-> Min" & vbNewLine & msgNoChiral
    msgMaxSingle = "Calculate the next Min(x,y) AND take the Rho-> MAX" & vbNewLine & msgNoChiral
    msgMaxMultiple = "1) Check IF each kNN is No Chiral !" & vbNewLine & "2) Take the Rho-> MAX"
    msgRhoCalc = "IF kNN = 1 It's a Direct case, take this node number as solution." & vbNewLine & "IF kNN > 1 You need to calculate Rho ..."
   
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    iCount = 0
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
            iCount = iCount + 1
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(29, 255, 131) 'strong green
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    
    'check if min for rho calculus:
    '2024/03/03
    'To do :
    
    If valMin = 1 And iCount = 1 Then
    'here:
        Call msgBoxTempFullCustom(msgMin, 7, gAppName, vbInformation)
'            Call AddShapeMin
            Call AddShapeTriangleDown
            
            ElseIf valMin = 1 And iCount > 1 Then
                Call msgBoxTempFullCustom(msgRhoCalc, 7, gAppName, vbExclamation)
                Call AddShapeTriangleDown
                
             ElseIf valMin <> 1 And iCount = 1 Then
                Call msgBoxTempFullCustom(msgMaxSingle, 7, gAppName, 64)
'                    Call AddShapeMax
                    Call AddShapeTriangleUp
                    
            ElseIf valMin <> 1 And iCount > 1 Then
                Call msgBoxTempFullCustom(msgMaxMultiple, 7, gAppName, 64)
'                    Call AddShapeMax
                    Call AddShapeTriangleUp
    End If
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub


Sub AddShapeMax()
'msoShapeUpArrow 35
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
'     Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeUpArrow, L, T, W, H)
    
    With oShapRectang
        With .Fill
            .ForeColor.RGB = RGB(0, 255, 0)
'            .BackColor.RGB = RGB(255, 0, 0)
'            .UserPicture sPathImage
        End With

        .Name = "arrowUpShape"
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
                    .Font.Size = 24
                    .Font.Color = vbYellow
                End With
        
    End With
End Sub


Sub AddShapeMin()
'msoShapeDownArrow 36
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
'     Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeDownArrow, L, T, W, H)
    
    With oShapRectang
        With .Fill
            .ForeColor.RGB = RGB(0, 255, 0)
'            .BackColor.RGB = RGB(255, 0, 0)
'            .UserPicture sPathImage
        End With

        .Name = "arrowUpShape"
        .TextFrame.Characters.Text = "Min rho"
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
                    .Font.Size = 24
                    .Font.Color = vbYellow
                End With
        
    End With
End Sub


Sub AddShapeTriangleUp()
'msoShapeIsoscelesTriangle 7
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
'     Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeIsoscelesTriangle, L, T, W, H)
    
    With oShapRectang
        .Name = "arrowUpShape"
'        .Rotation = 180
        With .Fill
            .ForeColor.RGB = RGB(0, 255, 0)
'            .BackColor.RGB = RGB(255, 0, 0)
'            .UserPicture sPathImage
        End With

        
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
                    .Font.Size = 24
                    .Font.Color = vbYellow
                End With
        
    End With
End Sub


Sub AddShapeTriangleDown()
'msoShapeIsoscelesTriangle 7
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
'     Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    Set oShapRectang = oWshapes.Shapes.AddShape(msoShapeIsoscelesTriangle, L, T, W, H)
    
    With oShapRectang
        .Name = "arrowDownShape"
        .Rotation = 180
        With .Fill
            .ForeColor.RGB = RGB(0, 255, 0)
'            .BackColor.RGB = RGB(255, 0, 0)
'            .UserPicture sPathImage
        End With

        
        .TextFrame.Characters.Text = "Min rho"
            With .Line
                .ForeColor.RGB = RGB(255, 0, 0)
                .Weight = 10
            End With
            With .TextFrame
                .AutoSize = True
                'Text orientation
                .Orientation = msoTextOrientationUpward
'                 .Orientation = msoTextOrientationVertical
            
            End With
                With .TextFrame.Characters
                    .Font.Bold = True
                    .Font.Size = 24
                    .Font.Color = vbYellow
                End With
        
    End With
End Sub
Sub deleteUpDownArrowShapes()
    Dim Sh As Shape
 
For Each Sh In ActiveSheet.Shapes
    If Sh.Name = "arrowUpShape" Or Sh.Name = "arrowDownShape" Then
        Sh.Delete
        Beep
    End If
Next
End Sub
Sub findAndLighValMinSecondOneLG_0062()
'alg_06_2
'Generic
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(103, 255, 216) 'light green
                
'                    Call noBorders
                    
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValMinSB_0063()
'alg_06_3
'Generic
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
    Dim customBlue1st$
    
    customBlue1st = RGB(0, 179, 242)
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = customBlue1st 'Strong blue
                Rng.Select
                'suited format for the Min:
                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValMinOneOneLB_0064()
'alg_06_4
'Generic
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
    Dim customBlue2nd$
    
    customBlue2nd = RGB(217, 245, 255)
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = customBlue2nd 'Light blue
                
'                    Call noBorders
                    
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub findAndLighValMinSecondTwoOrange_0065()
'alg_06_3
'Generic
    'Update cell in yellow
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMin!
    Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
      
    On Error Resume Next
    xTitleId = "Looking for min value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMin = Application.WorksheetFunction.Min(WorkRng)
    
    For Each Rng In WorkRng
        If Rng.Value = valMin Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = orangeColor
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub

Sub findAndLighValGradient_007()
'alg_07
'Generic
 'Gradient
'
  'Color in orange a new value choosed as min
    Dim Rng As Range
    Dim WorkRng As Range
      Dim trackedValue&
      Dim totalValues%
      Dim nbrUpdated%
      Dim remainingValues%
      
        Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
'    trackedValue = Application.InputBox("Searched value:", "Tracking")
    
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Select one value", "One cell", WorkRng.Address, Type:=8)
    trackedValue = WorkRng.Value
    
    'On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("DATA_RANGE", , WorkRng.Address, Type:=8)
    
    totalValues = 0
    nbrUpdated = 0
    For Each Rng In WorkRng
       
        If Rng.Value > 0 And Rng.Value <= CInt(trackedValue) Then
          
        Rng.Interior.Color = orangeColor
        nbrUpdated = nbrUpdated + 1
        End If
        
         totalValues = totalValues + 1
    Next
    
    remainingValues = totalValues - nbrUpdated
   
    MsgBox nbrUpdated & " items out of " & totalValues & " values have been updated." & vbNewLine & _
                            "Remaining values : " & remainingValues, vbInformation, gCR
                            
    Beep
    
                            
    If nbrUpdated < 2 Then
        MsgBox "=>DIRECT||==>", vbExclamation, "=>DIRECT||==>"
    
    End If
    Beep
    Beep
    
    
End Sub
Sub findAndLighValMax_008()
    'alg_08
    'Generic
    Dim Rng As Range
    Dim WorkRng As Range
    
    Dim xTitleId$
'    Dim valMin&
    Dim valMax As Single
      
    On Error Resume Next
    xTitleId = "Looking for MAX value"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    
    valMax = CSng(Application.WorksheetFunction.Max(WorkRng))
    
'    MsgBox valMax, vbInformation
    
    
    For Each Rng In WorkRng
        If Rng.Value = valMax Then
'            If Rng.Interior.Color = vbYellow Then
'                Rng.Interior.Color = vbGreen
'            End If
'                Rng.Value = ""
'                Rng.Interior.ColorIndex = 44
'                Rng.Interior.Color = RGB(0, 250, 0)
                Rng.Interior.Color = RGB(83, 210, 255) '#B5D2ED
                Rng.Select
                'suited format for the Min:
'                Call Mandatory
        End If
    Next
    Beep
    'Optimize
'    MsgBox "=> " & valMin, vbInformation, gAppName
    
End Sub
Sub DeleteAllShapes()
    Do Until ActiveSheet.Shapes.Count = 0
        ActiveSheet.Shapes(1).Delete
    Loop
End Sub
Sub Mandatory()

    With Selection.Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThick
    End With
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(153, 229, 255)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub ClearCellsRowContents()
'
    Selection.ClearContents
    ActiveCell.Select
End Sub
Sub ClearAll()
'! Clear data Matrix & ...
    Selection.Clear
    ActiveCell.Select
    
'! ...delete Shapes !
    Call deleteUpDownArrowShapes
    

End Sub
Sub CellBackgdOrange()
'

     Dim orangeColor$
    
    orangeColor = RGB(255, 192, 0)
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = orangeColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgdLightOrange()
'

     Dim orangeColor$
    
    orangeColor = RGB(255, 227, 171)
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = orangeColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgdRed()
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = vbRed
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgdYellow()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgdGray()
'
' CellBackgdYello Macro
' Cell's background yellow
'

     Dim grayColor$
    
    grayColor = RGB(174, 172, 172)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = grayColor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgdWhite()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = vbWhite
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub
Sub CellBackgRed()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub

'#TO DO

Sub sModulo3()
'Modulo 3

    Dim iDividend As Integer
    Dim iDivisor As Integer
    
'    MsgBox iDividend Math.mo
    

End Sub
Function Modulo3()
'Modulo3

End Function
Sub ongoingVertex()
' vertexPick Macro
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16711681
        .Size = 20
        .Bold = True
        .TintAndShade = 0
    End With
End Sub
Sub solutionVertex()
' solutionVertex Macro
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(51, 255, 143)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = vbBlack
        .Size = 22
        .Bold = True
        .TintAndShade = 0
    End With
    
    Call BordersGreen
    
End Sub
Sub BordersGreen()
'
' BordersGreen Macro
'    Range("C78").Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub BordersRed()
'
' BordersRed Macro

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    'Modif
      With Selection
        With .Font
        .Bold = True
        .Color = vbRed
       End With
    End With
    
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub openVerbatim_()
    'G:\IKRA\RECHERCHE\TSP\PUBLISH\TSP_N_14\NL_14\VerbatimAlgo.docx
    
End Sub

Sub openVerbatimEarlyBinding()
'! Bug à cause restriction de Microsoft ! (virus potentiel)
    Dim wordApp As Word.Application
    Dim strPath$
    strPath = "G:\IKRA\RECHERCHE\TSP\PUBLISH\TSP_N_14\NL_14\VerbatimAlgo.docx"
    Set wordApp = New Word.Application
     wordApp.Documents.Open strPath
    wordApp.Visible = True
End Sub
Sub openVerbatimLateBinding()
Dim wordApp As Object
    Dim strPath$
    strPath = "G:\IKRA\RECHERCHE\TSP\PUBLISH\TSP_N_14\NL_14\VerbatimAlgo.docx"
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open strPath
    wordApp.Visible = True
End Sub
Sub labelVertex()
'
' labelVertex Macro

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Color = -16711681
        .TintAndShade = 0
    End With
    With Selection.Font
        .Name = "Algerian"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16711681
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Size = 22
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
'/#TO DO
'##################################### subroutines ############################-
Sub noBorders()
'
' noborders Macro
'

'
'    Range("K87").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub CellBackgrdGreen()
'
' CellBackgdYello Macro
' Cell's background yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(51, 255, 143)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -11489280
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -11489280
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -11489280
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -11489280
        .TintAndShade = 0
        .Weight = xlThick
    End With
    ActiveCell.Select
End Sub


'#####################################  / ############################-


Sub TO_DO()
    'format display angle theta
End Sub

















