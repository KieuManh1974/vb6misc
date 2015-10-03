Attribute VB_Name = "Definition"
Option Explicit

Public moParser As ISaffronObject
Public msProgram As String

Sub main()
    InitialiseParser
    LoadProgram
    InitialiseCompiler
    Compiler.CompileProgram
End Sub

Public Sub InitialiseParser()
    Dim sDefinition As String
    
    sDefinition = Space$(FileLen(App.Path & "\arraylang.saf"))
    Open App.Path & "\arraylang.saf" For Binary As #1
    Get #1, , sDefinition
    Close #1
    
    If Not CreateRules(sDefinition) Then
        Debug.Print ErrorString
        MsgBox ErrorString
        End
    End If
    
    Set moParser = Rules("program")
End Sub

Public Sub LoadProgram()
    Dim sProgram As String
    
    msProgram = Space$(FileLen(App.Path & "\program.txt"))
    Open App.Path & "\program.txt" For Binary As #1
    Get #1, , msProgram
    Close #1
End Sub
