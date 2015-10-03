Attribute VB_Name = "Definition"
Option Explicit

Public oParser As IParseObject

Public Sub main()
    InitialiseParser
    CompileProgram
End Sub

Public Sub InitialiseParser()
    Dim Definition As String
    Dim oFSO As New FileSystemObject

    Definition = oFSO.OpenTextFile(App.Path & "\language2.pdl").ReadAll

    If Not SetNewDefinition(Definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = ParserObjects("program")
End Sub
