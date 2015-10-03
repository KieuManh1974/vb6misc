Attribute VB_Name = "Module1"
Option Explicit

Private oParse As IParseObject

Sub main()
    Definition
End Sub

Public Sub Definition()
    Dim sDef As String
    Dim oFS As New FileSystemObject
    
    sDef = oFS.OpenTextFile(App.Path & "/pdlsyntax.pdl").ReadAll
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
        
    Set oParse = ParserObjects("program")
End Sub
