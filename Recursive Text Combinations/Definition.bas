Attribute VB_Name = "Definition"
Option Explicit

Public oParser As IParseObject

Public Sub InitialiseParser()
    Dim sDef As String

    sDef = sDef & "level0 := AND invisible, (OR reverse1, reverse2), ?list_name, ?list_name, LIST level1, [','] MIN 0;"
    sDef = sDef & "level1 := AND invisible, reverse2, ?list_name, ?list_name, REPEAT level2 MIN 0;"
    sDef = sDef & "level2 := OR level3, identifier, text;"
    sDef = sDef & "level3 := AND ['{'], level0, ['}'];"
    sDef = sDef & "text := REPEAT OR (&['|'], '|'), (&['|'], '~'), (&['|'], ','), (&['|'], '{'), (&['|'], '}'), (&['|'], '$'), (&['|'], ':'), (&['|'], '@'), (&['|'], '#'), IN 0 TO 255, NOT '|', NOT '~', NOT '$', NOT '{', NOT '}', NOT ':', NOT ',', NOT '@', NOT '#' MIN 0;"
    sDef = sDef & "identifier := AND ['$'], index_name;"
    sDef = sDef & "index_name := {REPEAT IN CASE 'a' TO 'z', '0' TO '9'};"
    sDef = sDef & "list_name := AND identifier, [':'];"
    sDef = sDef & "invisible := OPTIONAL '@';"
    sDef = sDef & "reverse1 := OPTIONAL '~';"
    sDef = sDef & "reverse2 := OPTIONAL '#';"
    sDef = sDef & "reverse3 := OPTIONAL '^';"
        
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParser = ParserObjects("level0")
    
End Sub
