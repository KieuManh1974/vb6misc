Attribute VB_Name = "Program"
Option Explicit

Private ProgramContainer As New Container

Public Function CompileProgram()
    Dim oFSO As New FileSystemObject
    Dim oTree As New ParseTree

    Stream.Text = oFSO.OpenTextFile(App.Path & "\testprogram.txt", ForReading).ReadAll
    Set oTree = New ParseTree
    If Not oParser.Parse(oTree) Then
        MsgBox "Syntax error"
        End
    End If
    
    'ProgramBlock.CompileBlock oTree
        
    ProgramContainer.Compile oTree
End Function

