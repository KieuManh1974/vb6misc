Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    InitialiseNameParser
End Sub

Public Sub InitialiseNameParser()
    Dim Definition As String
    Dim oStripQuotes As IParseObject
    
    Definition = "strip_quotes := in 32, 123 to 124, 'x';"
                     
    If Not SetNewDefinition(Definition) Then
        Debug.Print ErrorString
        End
    End If
    
    Set oStripQuotes = ParserObjects("strip_quotes")
End Sub
