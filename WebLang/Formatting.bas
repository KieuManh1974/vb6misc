Attribute VB_Name = "Formatting"
Option Explicit

Private moParser As ISaffronObject

Private Enum Language
    laHTML
    laPHP
    laJavascript
End Enum

Private mvHTMLKeywords As Variant
Private mvPHPKeywords As Variant
Private mvJavascriptKeywords As Variant

Public Sub Main()
    IntialiseKeywords
    InitialiseParser
    ParseFiles
End Sub

Private Sub InitialiseParser()
    Dim sFile As String
    Dim sDefinition As String
    
    sFile = App.Path & "/WebLang.saf"
    sDefinition = String$(FileLen(sFile), Chr(0))
    Open sFile For Binary As #1
    Get #1, , sDefinition
    Close #1
    
    If Not CreateRules(sDefinition) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set moParser = Rules("weblang")
    
'    Set moParser = Rules("text")
'    SaffronStream.Text = "`test`"
'    Dim oTest As SaffronTree
'    Set oTest = New SaffronTree
'    If moParser.Parse(oTest) Then
'        Stop
'    Else
'        Stop
'    End If
End Sub

Private Sub ParseFiles()
    Dim sFile As String
    Dim sContents As String
    Dim sFormattedFile As String
    Dim vExtensions As Variant
    Dim vExtension As Variant
    vExtensions = Array("txt")
    
    For Each vExtension In vExtensions
        sFile = Dir(App.Path & "\*." & vExtension)
        On Error Resume Next
        MkDir App.Path & "\_Converted"
        On Error GoTo 0
        While Not sFile = ""
            sContents = String$(FileLen(sFile), Chr(0))
            Open sFile For Binary As #1
            Get #1, , sContents
            Close #1
        
            sFormattedFile = ParseFile(sContents)
            On Error Resume Next
            Kill App.Path & "\_Converted\" & sFile
            On Error GoTo 0
            Open App.Path & "\_Converted\" & sFile For Binary As #1
            Put #1, , sFormattedFile
            Close #1
            sFile = Dir
        Wend
    Next
    MsgBox "Files have been formatted.", vbOKOnly
End Sub

Private Function ParseFile(sContents As String) As String
    Dim oTree As New SaffronTree
    
    SaffronStream.Text = sContents
    
    If moParser.Parse(oTree) Then
        ParseFile = FormatWebLang(oTree)
    Else
        ParseFile = sContents
    End If
    
    'Debug.Print ParseFile
End Function


Private Function FormatWebLang(oTree As SaffronTree) As String
    FormatWebLang = FormatStructList(oTree(2), 0)
End Function

Private Function FormatStructList(oTree As SaffronTree, ByVal lIndentLevel As Long, Optional ByVal bMultiline As Boolean = True) As String
    Dim oLine As SaffronTree
    Dim bText As Boolean
    Dim bDelimiter As Boolean
    
    For Each oLine In oTree.SubTree
        If Not bDelimiter Then
            bText = oLine.Index = 1
            If bText Then
                FormatStructList = FormatStructList & FormatText(oLine(1), lIndentLevel, bMultiline)
            Else
                FormatStructList = FormatStructList & FormatStruct(oLine(1), lIndentLevel, bMultiline)
            End If
        End If
        bDelimiter = Not bDelimiter
    Next
End Function

Private Function FormatText(oTree As SaffronTree, ByVal lIndentLevel As Long, Optional ByVal bMultiline As Boolean = True) As String
    FormatText = oTree.Text
    FormatText = Replace$(FormatText, "\`", "`")
    FormatText = Replace$(FormatText, "\\", "\")
    FormatText = Indent(lIndentLevel, bMultiline) & FormatText & NextLine(bMultiline)
End Function

Private Function FormatStruct(oTree As SaffronTree, ByVal lIndentLevel As Long, Optional ByVal bMultiline As Boolean = True) As String
    Dim oLine As SaffronTree
    Dim sIdentifier As String
    Dim vParameters As Variant
    Dim bHasParameters As Boolean
    Dim bHasBlock As Boolean
    
    sIdentifier = oTree(1).Text
    bHasParameters = oTree(3).Index = 1
    bHasBlock = oTree(5).Index = 1
    
    If bHasParameters Then
        FormatStruct = Indent(lIndentLevel, bMultiline) & OpenTag(sIdentifier, FormatParameters(oTree(3)(1)(2), lIndentLevel + 1)) & NextLine(bMultiline)
    Else
        FormatStruct = Indent(lIndentLevel, bMultiline) & OpenTag(sIdentifier) & NextLine(bMultiline)
    End If
    
    If bHasBlock Then
        FormatStruct = FormatStruct & FormatStructList(oTree(5)(1)(2), lIndentLevel + 1, bMultiline)
    End If
    FormatStruct = FormatStruct & Indent(lIndentLevel, bMultiline) & CloseTag(sIdentifier) & NextLine(bMultiline)
End Function

' Format for HTML attributes
Private Function FormatParameters(oTree As SaffronTree, ByVal lIndentLevel As Long) As String
    Dim oLine As SaffronTree
    Dim lCount As Long
    Dim sIdentifier As String
    Dim bHasValue As Boolean
    Dim bHasIdentifier As Boolean
    Dim bDelimiter As Boolean
    
    For Each oLine In oTree.SubTree
        If Not bDelimiter Then
            FormatParameters = FormatParameters & " "
            bHasIdentifier = oLine(1).Index
            If bHasIdentifier Then
                FormatParameters = FormatParameters & oLine(1)(1)(1).Text & " = """
            End If
            
            Select Case oLine(3).Index
                Case 1
                    FormatParameters = FormatParameters & FormatText(oLine(3)(1), 0, False)
                Case 2
                    FormatParameters = FormatParameters & FormatStructList(oLine(3)(1)(2), 0, False)
            End Select
            
            If bHasIdentifier Then
                FormatParameters = FormatParameters & """"
            End If
        End If
        bDelimiter = Not bDelimiter
    Next
End Function

Private Function OpenTag(ByVal sTagName As String, Optional ByVal sParameters As String) As String
    OpenTag = "<" & sTagName & sParameters & ">"
End Function

Private Function CloseTag(ByVal sTagName As String) As String
    CloseTag = "</" & sTagName & ">"
End Function

Private Function Indent(ByVal lIndentLevel As Long, Optional bMultiline As Boolean = True) As String
    If bMultiline Then Indent = String$(lIndentLevel * 4, " ")
End Function

Private Function IndentChild(ByVal lIndentLevel As Long, Optional bMultiline As Boolean = True) As String
    If bMultiline Then Indent = String$(lIndentLevel * 4 + 4, " ")
End Function

Private Function NextLine(ByVal bMultiline As Boolean) As String
    If bMultiline Then NextLine = vbCrLf
End Function

Private Function LanguageSwitch(sKeyword As String) As Language
    Dim lIndex As Long

    For lIndex = 0 To UBound(mvHTMLKeywords)
        If mvHTMLKeywords(lIndex) = LCase$(sKeyword) Then
            LanguageSwitch = laHTML
            Exit Function
        End If
    Next

    For lIndex = 0 To UBound(mvPHPKeywords)
        If mvPHPKeywords(lIndex) = LCase$(sKeyword) Then
            LanguageSwitch = laPHP
            Exit Function
        End If
    Next

    For lIndex = 0 To UBound(mvJavascriptKeywords)
        If mvJavascriptKeywords(lIndex) = LCase$(sKeyword) Then
            LanguageSwitch = laJavascript
            Exit Function
        End If
    Next
End Function

Private Sub InitialiseKeywords()
    mvHTMLKeywords = Array("html", "head", "body", "title", "div", "span", "table", "tr", "th", "td")
    mvPHPKeywords = Array("_for", "_switch", "_while")
    mvJavascriptKeywords = Array("for", "switch", "case", "while")
End Sub
