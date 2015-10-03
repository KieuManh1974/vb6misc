Attribute VB_Name = "Definition"
Option Explicit

Public oParsePosition As IParseObject
Public oParseRelationship As IParseObject
Public oParseLine As IParseObject

Public Const pi2  As Double = 6.28318530717959

Public Sub DoDefinition()
    Dim sDef As String
    
    sDef = sDef & "pos := REPEAT IN '0' TO '9', '.', '-';"
    sDef = sDef & "name := REPEAT IN 32 TO 255, NOT '|' MIN 0;"
    sDef = sDef & "position := AND ['P:'], reference, ['|'], name, ['|'], pos, ['|'], pos, ['|'], pos, ['|'], pos;"
    sDef = sDef & "reference := REPEAT IN CASE 'A' TO 'F', '0' TO '9';"
    sDef = sDef & "relationship := AND ['R:'], reference, ['|'], reference, ['|'], pos, ['|'], pos;"
    sDef = sDef & "offset := AND ['O:'], pos, ['|'], pos;"
    sDef = sDef & "zoom := AND ['Z:'], pos;"
    sDef = sDef & "line := OR position, relationship, offset, zoom;"
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParsePosition = ParserObjects("position")
    Set oParseRelationship = ParserObjects("relationship")
    Set oParseLine = ParserObjects("line")
End Sub



