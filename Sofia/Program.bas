Attribute VB_Name = "Program"
Option Explicit

Private ProgramAnalyse As New Analyse
Private ProgramCompose As New Compose
Private ProgramGenerate As New GenerateIntermediateCode

Public Function CompileProgram()
    Dim oFSO As New FileSystemObject
    Dim oTree As New SaffronTree
    Dim oAnalysed As New clsNode
    Dim oBlocks As Block
    
    SaffronStream.Text = oFSO.OpenTextFile(App.Path & "\testprogram.txt", ForReading).ReadAll
    Set oTree = New SaffronTree
    If Not oParser.Parse(oTree) Then
        MsgBox "Syntax error"
        End
    End If
    
    ProgramAnalyse.CompileParseTree oTree, oAnalysed
    Set oBlocks = ProgramCompose.CompileBlocks(oAnalysed)
    oBlocks.msName = "main"
    ProgramGenerate.CompileAssembly oBlocks
End Function

