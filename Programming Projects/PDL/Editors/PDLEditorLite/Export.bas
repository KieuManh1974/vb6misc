Attribute VB_Name = "Export"
Option Explicit

Public Sub ExportToVisualBasic(sFile As String, sPDLFile As String)
    Dim ts As TextStream
    Dim lindex As Long
    Dim sLineText As String
    
    With New FileSystemObject
        Set ts = .CreateTextFile(sFile, True)
        ts.WriteLine "Option Explicit"
        ts.WriteLine
        ts.WriteLine "Public oParser As IParseObject"
        ts.WriteLine
        ts.WriteLine "Public Sub InitialiseParser()"
        ts.WriteLine "    Dim definition As String"
        ts.WriteLine "    Dim oFSO as New FileSystemObject"
        ts.WriteLine
        ts.WriteLine "    definition=oFSO.OpenTextFile(""" & sPDLFile & """).ReadAll"
        ts.WriteLine
        ts.WriteLine "    If Not SetNewDefinition(definition) Then"
        ts.WriteLine "        Debug.Print ErrorString"
        ts.WriteLine "        End"
        ts.WriteLine "    End If" & vbCrLf
        ts.WriteLine
        ts.WriteLine "    Set oParser = ParserObjects("""")"
        ts.WriteLine "End Sub"
    End With
End Sub

Public Sub KeywordAnd(sDefinitionName)

End Sub

Public Sub KeywordOr(sDefinitionName)
' Private Function sDefinitionName(
' Select Case oParseTree.Index
' End Select
End Sub

Public Sub KeywordNot(sDefinitionName)

End Sub

Public Sub KeywordRepeat(sDefinitionName)

End Sub

Public Sub KeywordList(sDefinitionName)

End Sub

Public Sub KeywordIn(sDefinitionName)

End Sub

Public Sub KeywordEach(sDefinitionName)

End Sub

Public Sub KeywordOptional(sDefinitionName)

End Sub

Public Sub KeywordLiteral(sDefinitionName)

End Sub

