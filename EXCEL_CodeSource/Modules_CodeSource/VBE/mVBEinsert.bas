Attribute VB_Name = "mVBEinsert"
Option Explicit
 Sub AddProcedureToModule()
 'ajoute une simple procédure "Hello World" nommée SayHello à la fin du module nommé Module1.
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim lineNum As Long
        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("Module1")
        Set CodeMod = VBComp.CodeModule
        
        With CodeMod
            lineNum = .CountOfLines + 1
            .InsertLines lineNum, "Public Sub SaySaleme()"
            lineNum = lineNum + 1
            .InsertLines lineNum, "    MsgBox " & DQUOTE & "Hello World" & DQUOTE
            lineNum = lineNum + 1
            .InsertLines lineNum, "End Sub"
        End With
    End Sub
Sub InsertInProcedure()
 'ajoute une simple procédure "Hello World" nommée SayHello à la fin du module nommé Module1.
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim lineNum As Long
        Dim VBCompo As VBComponent

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("Module1")
'        Set VBCompo = VBComp.Type
        
        
        
        
        With VBCompo.CodeModule
            lineNum = 3
            .InsertLines lineNum, "Say Salem !"
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "    MsgBox ""Hello World"""
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "End Sub"
        End With
    End Sub
Sub codeInsert(nomModule$, nomMacro$, startLine%)
'
'Dim nomModule$
'Dim nomMacro$
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(nomModule)
    Dim startProcLine%
    
    'data To insert:
    Dim myDate As Date
    myDate = Date$
            
    With VBComp.CodeModule
        startProcLine = .ProcStartLine(nomMacro, vbext_pk_Proc)
'        startLine = .ProcStartLine(nomMacro, vbext_pk_Proc)
        .InsertLines startProcLine + startLine, "'Date : " & vbNewLine & myDate
    End With
End Sub

Sub test_codeInsert()
Dim nomModule$
Dim nomMacro$
Dim lineNum%

nomModule = "Module1"
nomMacro = "testInsert"
lineNum = 5

    codeInsert nomModule, nomMacro, lineNum
    
    
End Sub
 Sub AddProcedureToModuleModif()
 'ajoute une simple procédure "Hello World" nommée SayHello à la fin du module nommé Module1.
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim lineNum As Long
'        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("Module1")
        Set CodeMod = VBComp.CodeModule
        
        With CodeMod
            lineNum = .CountOfLines + 1
            .InsertLines lineNum, "Public Sub SaySalem3()"
            lineNum = lineNum + 1
            .InsertLines lineNum, "    MsgBox ""Salem !"""
            lineNum = lineNum + 1
            .InsertLines lineNum, "End Sub"
        End With
    End Sub

