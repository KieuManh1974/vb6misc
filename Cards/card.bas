Attribute VB_Name = "Module2"
Option Explicit

Public oParser As IParseObject

Public Sub InitialiseParser()
    Dim definition As String
    Dim oFSO As New FileSystemObject

    definition = oFSO.OpenTextFile(App.Path & "\Cards.pdl").ReadAll

    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = ParserObjects("hand")
End Sub
