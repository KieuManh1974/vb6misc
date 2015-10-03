Attribute VB_Name = "FileNameManage"
Option Explicit

Public oStripQuotes As IParseObject


Public Sub InitialiseNameParser()
    Dim Definition As String

    Definition = "strip_quotes := {AND [?'""'], REPEAT IN 32 TO 255 UNTIL (OR EOS, '""')};"
                     
    If Not SetNewDefinition(Definition) Then
        Debug.Print ErrorString
        End
    End If
    
    Set oStripQuotes = ParserObjects("strip_quotes")
End Sub


