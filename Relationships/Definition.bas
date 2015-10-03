Attribute VB_Name = "Definition"
Option Explicit

Public oParsePosition As IParseObject
Public oParseRelationship As IParseObject

Public Positions As PositionList
Public Relations As RelationshipList
Public FileIOs As FileIO

Public Sub DoDefinition()
    Dim sDef As String
    
    sDef = sDef & "pos := REPEAT IN '0' TO '9', '.', '-';"
    sDef = sDef & "name := REPEAT IN 32 TO 255, NOT '|';"
    sDef = sDef & "position := AND ['P:'], reference, ['|'], name, ['|'], pos, ['|'], pos;"
    sDef = sDef & "reference := REPEAT IN CASE 'A' TO 'F', '0' TO '9';"
    sDef = sDef & "relationship := AND ['R:'], reference, ['|'], reference;"
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParsePosition = ParserObjects("position")
    Set oParseRelationship = ParserObjects("relationship")
End Sub

Public Sub Initialise(oCanvasRef As Form)
    Set Positions = New PositionList
    Set Relations = New RelationshipList
    
    Set FileIOs = New FileIO
    Set FileIOs.CanvasRef = oCanvasRef
    Set FileIOs.PositionsRef = Positions
    Set FileIOs.RelationsRef = Relations
    FileIOs.FileStore = App.Path & "\diagram.txt"
    FileIOs.ReadFile
End Sub

Public Sub DeInitialise()
    FileIOs.WriteFile
End Sub

