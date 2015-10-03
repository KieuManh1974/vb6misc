Attribute VB_Name = "ErrorTest"
Sub main()
    Dim oTest As IParseObject
    Dim oTree As New ParseTree
    
    InitialiseParser
    
    Set oTest = ParserObjects("test")
    
    ParserTextString.ParserText = "B"
    
    If Not oTest.Parse(oTree) Then
        Debug.Print "failsed"
    End If
End Sub

Public Sub InitialiseParser()
    Dim definition As String

    definition = "test := LIST 'A','|';"
    
    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If

End Sub
