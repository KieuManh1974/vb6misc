Attribute VB_Name = "Module1"
Public oParser As IParseObject

Sub Main()
    InitialiseParser
End Sub

Private Sub InitialiseParser()
    Dim definition As String
    Dim FSO As New FileSystemObject
    
    definition = FSO.OpenTextFile(App.Path & "\NumberCruncher.pdl").ReadAll
    
    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        Exit Sub
    End If
    Set oParser = ParserObjects("numbers")
End Sub
