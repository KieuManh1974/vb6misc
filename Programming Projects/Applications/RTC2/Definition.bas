Attribute VB_Name = "Definition"
Option Explicit

Public oParser As IParseObject
Public AllSets As New Dictionary

Public Sub InitialiseParser()
    Dim oFS As New FileSystemObject
    
    Dim sDef As String

    sDef = oFS.OpenTextFile(App.Path & "\rtc2.pdl").ReadAll
            
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParser = ParserObjects("set")
End Sub
