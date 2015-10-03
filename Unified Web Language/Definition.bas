Attribute VB_Name = "Definition"
Option Explicit

Public goUWLParser As ISaffronObject

Public Sub Main()
    InitialiseParser
    ParseFiles
End Sub

Private Sub InitialiseParser()
    Dim sFile As String
    Dim sDefinition As String
    
    sFile = App.Path & "/UnifiedWebLanguage.saf"
    sDefinition = String$(FileLen(sFile), Chr(0))
    Open sFile For Binary As #1
    Get #1, , sDefinition
    Close #1
    
    If Not CreateRules(sDefinition) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set goUWLParser = Rules("contents")
End Sub

Private Sub ParseFiles()
    Dim sFile As String
    Dim sContents As String
    Dim sCompiledFile As String
    Dim vExtensions As Variant
    Dim vExtension As Variant
    Dim sPath As String
    
    vExtensions = Array("uwl")
    
    For Each vExtension In vExtensions
        sPath = App.Path
        sFile = Dir(sPath & "\*." & vExtension)
        On Error Resume Next
        MkDir App.Path & "\compiled"
        On Error GoTo 0
        While Not sFile = ""
            sContents = String$(FileLen(sPath & "\" & sFile), Chr(0))
            Open sFile For Binary As #1
            Get #1, , sContents
            Close #1
        
            sCompiledFile = ParseFile(sContents)
            On Error Resume Next
            Kill App.Path & "\compiled\" & sFile
            On Error GoTo 0
            Open App.Path & "\compiled\" & sFile For Binary As #1
            Put #1, , sCompiledFile
            Close #1
            sFile = Dir
        Wend
    Next
    MsgBox "Files have been formatted.", vbOKOnly
End Sub

Private Function ParseFile(sContents As String) As String
    Dim oTree As New SaffronTree
    
    SaffronStream.Text = sContents
    
    If goUWLParser.Parse(oTree) Then
        ParseFile = CompileUWL(oTree, 0)
    Else
        ParseFile = sContents
    End If
    
    'Debug.Print ParseFile
End Function

Private Function CompileUWL(oTree As SaffronTree, ByVal lIndentLevel As Long) As String
    Dim sName As String
    Dim sOutput As String
    Dim sTag As String
    Dim oContent As SaffronTree
    
    sTag = oTree(1).Text
    
    sOutput = Indent(lIndentLevel) & "<" & sTag & ">" & NewLine

    For Each oContent In oTree(2).SubTree
        sOutput = sOutput & CompileUWL(oContent, lIndentLevel + 1)
    Next

    sOutput = sOutput & Indent(lIndentLevel) & "</" & sTag & ">" & NewLine
    
    CompileUWL = sOutput
End Function

Private Function Indent(ByVal lLevel As Long) As String
    Indent = String$(lLevel * 4, " ")
End Function

Private Function NewLine()
    NewLine = vbCrLf
End Function
