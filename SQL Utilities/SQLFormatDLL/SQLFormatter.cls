VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Parsers
Private oSQLParser As IParseObject
Private oVBCodeParser As IParseObject

Private oFieldAlias As IParseObject

' Options
Public KeywordStyle As Integer
Public OperatorStyle As Integer
Public FunctionStyle As Integer
Public FieldStyle As Integer
Public TableStyle As Integer
Public QuoteStyle As Integer
Public ASSelectStyle As Integer
Public ASFromStyle As Integer
Public WHEREExpression As Integer
Public VisualBasicCode As Integer
Public TableAliasing As Integer
Public VBSQLIndex As Integer
Public Indented As Integer
Public Style As Integer
Public IndentedSubQuery As Integer

Public OptionsChanged As Boolean

' Format
Private sIndent As String

Public colTableAliases As Dictionary
Public colTableUnaliases As Dictionary
Public colTables As Dictionary

Private sSpecialCharacters As String

Private mbIndented As Boolean
Private msNewLine As String

Private Sub Class_Initialize()
    InitialiseParser
    
    KeywordStyle = 3
    OperatorStyle = 3
    FunctionStyle = 3
    FieldStyle = 3
    TableStyle = 3
    QuoteStyle = 2
    WHEREExpression = 0
    TableAliasing = 3
    ASSelectStyle = 1
    ASFromStyle = 1
    VBSQLIndex = 0
    Indented = 0
    Style = 0
    IndentedSubQuery = 0
End Sub

Public Sub InitialiseParser()
    Dim sDef As String
    Dim oTS As TextStream
    Dim oFSO As New FileSystemObject
    
    sDef = oFSO.OpenTextFile(App.Path & "\SQLViewer.pdl").ReadAll
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
    End If
        
    Set oSQLParser = ParserObjects("sql_statement")
    Set oVBCodeParser = ParserObjects("vb_sql")
End Sub

Public Function FormatSQL(ByVal sSQL As String, bOK As Boolean, lErrorPos As Long) As String
    Dim oTree As ParseTree
    Dim sFirstSQL As String
    Dim sSQLPure As String
    Dim oLine As ParseTree
    Dim oSQLPart As ParseTree
    Dim lCount As Long
    
    mbIndented = Indented = 1
    msNewLine = IIf(mbIndented, vbCrLf, "")
    
    InitialiseSpecialCharacters
    
    Stream.Text = sSQL
    Set oTree = New ParseTree
    
    If oVBCodeParser.Parse(oTree) Then
        If Stream.Position > Len(sSQL) Then
            For Each oLine In oTree.SubTree
                For Each oSQLPart In oLine.SubTree
                    sSQLPure = sSQLPure & oSQLPart.Text
                Next
            Next
        Else
            bOK = False
            lErrorPos = Stream.Position
            FormatSQL = sSQL
            Exit Function
        End If
    Else
        sSQLPure = sSQL
    End If
    
    Stream.Text = sSQLPure
    Set oTree = New ParseTree
    
    If oSQLParser.Parse(oTree) Then
        Set colTableAliases = New Dictionary
        Set colTableUnaliases = New Dictionary
        Set colTables = New Dictionary
        
        sIndent = IIf(Indented = 1, "    ", " ")
        
        FormatSQL = FormatStatement(oTree)

        For lCount = 1 To 1
        'do
            If Stream.Position <= Len(Stream.Text) Then
                lErrorPos = Stream.Position
                FormatSQL = sSQLPure
                bOK = False
                Exit Function
            End If
            
'            sSQL = FormatSQL
'            Stream.Text = sSQL
'            Set oTree = New ParseTree
'            oSQLParser.Parse oTree
            FormatSQL = FormatStatement(oTree)
        'Loop Until FormatSQL = sSQL
        Next

        If VisualBasicCode = 0 Then
            FormatSQL = FormatForVisualBasic(FormatSQL)
        End If
        bOK = True
    Else
        lErrorPos = Stream.Position
        FormatSQL = sSQL
        bOK = False
    End If
End Function

Private Function FormatForVisualBasic(ByVal sSQL As String) As String
    Dim vLines As Variant
    Dim vLine As Variant
    Dim sSQLIndex As String
    
    If VBSQLIndex > 0 Then
        sSQLIndex = VBSQLIndex
    End If
    
    vLines = Split(sSQL, vbCrLf)
    
    FormatForVisualBasic = "sSQL" & sSQLIndex & " = """"" & vbCrLf
    
    For Each vLine In vLines
        FormatForVisualBasic = FormatForVisualBasic & "sSQL" & sSQLIndex & " = sSQL" & sSQLIndex & " & "" " & vLine & """" & vbCrLf
    Next
End Function

Private Function FormatStatement(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1
            FormatStatement = FormatShapeCommand(oTree(1))
        Case 2
            FormatStatement = FormatSQLStatement(oTree(1))
    End Select
End Function

Private Function FormatShapeCommand(oTree As ParseTree, Optional ByVal lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)

    FormatShapeCommand = FormatShapeClause(oTree(1), lIndentationDepth) & msNewLine & FormatAppendClause(oTree(2), lIndentationDepth)
End Function

Private Function FormatShapeClause(oTree As ParseTree, ByVal lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
    
    FormatShapeClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    FormatShapeClause = FormatShapeClause & sIndentation1 & "{" & StripNewLine(StripIndentation(FormatSelectStatement(oTree(2), lIndentationDepth + 1), lIndentationDepth + 1)) & "}"
    If oTree(3).Index = 1 Then
        FormatShapeClause = FormatShapeClause & FormatAsClause(oTree(3)(1))
    End If
End Function

Private Function FormatAppendClause(oTree As ParseTree, ByVal lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
    
    FormatAppendClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    FormatAppendClause = FormatAppendClause & FormatAppendList(oTree(2), lIndentationDepth + 1)
End Function

Private Function FormatAppendList(oTree As ParseTree, ByVal lIndentationDepth As Long) As String
    Dim oAppendItem As ParseTree
    
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
        
    For Each oAppendItem In oTree.SubTree
        FormatAppendList = FormatAppendList & sIndentation & "("
        Select Case oAppendItem(1).Index
            Case 1 ' ChildClause
                FormatAppendList = FormatAppendList & StripIndentation(FormatChildClause(oAppendItem(1)(1), lIndentationDepth), lIndentationDepth)
            Case 2 ' ShapeStatement
                FormatAppendList = FormatAppendList & "(" & StripIndentation(FormatShapeCommand(oAppendItem(1)(1)(1), lIndentationDepth), lIndentationDepth) & ")"
        End Select
        FormatAppendList = FormatAppendList & msNewLine & sIndentation & FormatRelateClause(oAppendItem(2))
        FormatAppendList = FormatAppendList & ")," & msNewLine
    Next
    FormatAppendList = Left$(FormatAppendList, Len(FormatAppendList) - 3)
End Function

Private Function FormatChildClause(oTree As ParseTree, ByVal lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)

    FormatChildClause = sIndentation & "{" & StripNewLine(StripIndentation(FormatSelectStatement(oTree(1), lIndentationDepth), lIndentationDepth)) & "}"
    If oTree(2).Index = 1 Then
        FormatChildClause = FormatChildClause & FormatAsClause(oTree(2)(1))
    End If
End Function

Private Function FormatRelateClause(oTree As ParseTree) As String
    Dim oRelateItem As ParseTree
    
    FormatRelateClause = ChangeCase(oTree(1).Text, KeywordStyle) & " "
    
    For Each oRelateItem In oTree(2).SubTree
        FormatRelateClause = FormatRelateClause & oRelateItem(1).Text & " " & ChangeCase(oRelateItem(2).Text, KeywordStyle) & " " & oRelateItem(3).Text & ", "
    Next
    FormatRelateClause = Left$(FormatRelateClause, Len(FormatRelateClause) - 2)
End Function

Private Function FormatAsClause(oTree As ParseTree) As String
    FormatAsClause = " " & ChangeCase(oTree(1).Text, KeywordStyle) & " " & oTree(2).Text
End Function

Private Function FormatSQLStatement(oTree As ParseTree) As String
    If oTree(1).Index = 1 Then
        FormatSQLStatement = FormatParametersClause(oTree(1)(1))
    End If
    
    Select Case oTree(2).Index
        Case 1 'UNION (SELECT)
            FormatSQLStatement = FormatSQLStatement & StripNewLine(FormatUnionStatement(oTree(2)(1)))
        Case 2 ' UPDATE
            FormatSQLStatement = FormatSQLStatement & StripNewLine(FormatUpdateStatement(oTree(2)(1)))
        Case 3 ' INSERT
            FormatSQLStatement = FormatSQLStatement & StripNewLine(FormatInsertReplaceStatement(oTree(2)(1)))
        Case 4 ' DELETE
            FormatSQLStatement = FormatSQLStatement & StripNewLine(FormatDeleteStatement(oTree(2)(1)))
    End Select
    
    If Indented = 0 Then
        FormatSQLStatement = Mid$(FormatSQLStatement, 2)
    End If
End Function

Private Function FormatParametersClause(oTree As ParseTree) As String
    Dim oParameter As ParseTree
    Dim sParameterList As String
    
    FormatParametersClause = ChangeCase(oTree(1).Text, KeywordStyle) & IIf(Indented = 1, vbCrLf, "")
    For Each oParameter In oTree(2).SubTree
        sParameterList = sParameterList & "," & IIf(Indented = 1, vbCrLf, "") & "    " & ChangeCase(oParameter(1).Text, FieldStyle) & " " & oParameter(2).Text & oParameter(3).Text
    Next
    FormatParametersClause = FormatParametersClause & Mid$(sParameterList, 4) & ";" & IIf(Indented = 1, vbCrLf, "")
End Function

Private Function FormatUpdateStatement(oTree As ParseTree) As String
    Dim sIndentation As String
    
    sIndentation = Indentation(0)
    
    FormatUpdateStatement = sIndentation & FormatUpdateClause(oTree(1)) & msNewLine
    FormatUpdateStatement = FormatUpdateStatement & sIndentation & FormatSetClause(oTree(2)) & msNewLine
    If oTree(3).Index = 1 Then
        FormatUpdateStatement = FormatUpdateStatement & FormatWhereClause(oTree(3)(1), 0)
    End If
End Function

Private Function FormatUpdateClause(oTree As ParseTree) As String
    FormatUpdateClause = FormatUpdateClause & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine & Indentation(1) & FormatTable(oTree(2))
End Function

Private Function FormatSetClause(oTree As ParseTree) As String
    Dim oFieldValue As ParseTree
    Dim sSetFields As String
    
    FormatSetClause = FormatSetClause & ChangeCase(oTree(1).Text, KeywordStyle)
    
    For Each oFieldValue In oTree(2).SubTree
        sSetFields = sSetFields & "," & IIf(Indented = 1, vbCrLf, "") & sIndent & FormatQualifiedField(oFieldValue(1)) & " = " & FormatLevelX(1, oFieldValue(3))
    Next
    FormatSetClause = FormatSetClause & Mid$(sSetFields, 2)
End Function

Private Function FormatInsertReplaceStatement(oTree As ParseTree) As String
    FormatInsertReplaceStatement = FormatInsertReplaceIntoClause(oTree(1))
    
    If oTree(2).Index = 1 Then
        FormatInsertReplaceStatement = FormatInsertReplaceStatement & FormatValueList(oTree(2)(1)) & msNewLine
    Else
        FormatInsertReplaceStatement = FormatInsertReplaceStatement & FormatSelectStatement(oTree(2)(1), 1, False)
    End If
End Function

Private Function FormatInsertReplaceIntoClause(oTree As ParseTree) As String
    Dim oField As ParseTree
    Dim sFields As String
    Dim sIndentation As String
    Dim sIndentation1 As String
        
    sIndentation = Indentation(0)
    sIndentation1 = Indentation(1)
        
    FormatInsertReplaceIntoClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & " " & ChangeCase(oTree(2).Text, KeywordStyle) & msNewLine & sIndentation1 & ChangeCase(oTree(3).Text, TableStyle) & msNewLine & sIndentation1
    For Each oField In oTree(4)(1).SubTree
        sFields = sFields & "," & msNewLine & sIndentation1 & oField.Text
    Next
    FormatInsertReplaceIntoClause = FormatInsertReplaceIntoClause & "(" & StripIndentation(Mid$(sFields, 1 + Len("," & msNewLine)), 1) & ")" & msNewLine
End Function

Private Function FormatValueList(oTree As ParseTree) As String
    Dim oField As ParseTree
    Dim sFields As String
    
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(0)
    sIndentation1 = Indentation(1)
        
    FormatValueList = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine & sIndentation1
    For Each oField In oTree(2).SubTree
        sFields = sFields & "," & msNewLine & sIndentation1 & FormatLevelX(1, oField)
    Next
    FormatValueList = FormatValueList & "(" & StripIndentation(Mid$(sFields, 1 + Len("," & msNewLine)), 1) & ")"
End Function

Private Function FormatUnionStatement(oTree As ParseTree, Optional ByVal lIndentationDepth As Long, Optional ByVal bBrackets As Boolean) As String
    Dim oSelect As ParseTree
    Dim sIndentation As String
    Dim sSelectStatement As String
    
    sIndentation = Indentation(lIndentationDepth)
    
    For Each oSelect In oTree.SubTree
        If UCase$(oSelect(1).Text) = "UNION" Then
            FormatUnionStatement = FormatUnionStatement & sIndentation & ChangeCase(oSelect(1).Text, KeywordStyle)
            If oSelect(2).Index = 1 Then
                FormatUnionStatement = FormatUnionStatement & " " & ChangeCase(oSelect(2)(1).Text, KeywordStyle)
            End If
            FormatUnionStatement = FormatUnionStatement & msNewLine
        Else
            FormatUnionStatement = FormatUnionStatement & FormatSelectStatement(oSelect(1)(1), lIndentationDepth)
        End If
    Next
    If bBrackets Then
        FormatUnionStatement = sIndentation & "(" & StripNewLine(StripIndentation(FormatUnionStatement, lIndentationDepth)) & ")" & msNewLine
    End If
End Function

Private Function FormatSelectStatement(oTree As ParseTree, ByVal lIndentationDepth As Long, Optional ByVal bBrackets As Boolean) As String
    Dim sIndentation As String
    
    sIndentation = Indentation(lIndentationDepth)
    
    If oTree(1).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatWithClause(oTree(1), lIndentationDepth)
    End If
    
    FormatSelectStatement = FormatSelectStatement & FormatSelectClause(oTree(2), lIndentationDepth)
    
    If oTree(3).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatIntoClause(oTree(3)(1), lIndentationDepth)
    End If
    
    If oTree(4).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatFromClause(oTree(4)(1), lIndentationDepth)
    End If

    If oTree(5).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatWhereClause(oTree(5)(1), lIndentationDepth)
    End If
    
    If oTree(6).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatGroupByClause(oTree(6)(1), lIndentationDepth)
    End If
    
    If oTree(7).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatHavingClause(oTree(7)(1), lIndentationDepth)
    End If
    
    If oTree(8).Index = 1 Then
        FormatSelectStatement = FormatSelectStatement & FormatOrderByClause(oTree(8)(1), lIndentationDepth)
    End If
End Function

Private Function FormatDeleteStatement(oTree As ParseTree) As String
    Dim sFrom As String
    
    FormatDeleteStatement = Indentation(0) & FormatDeleteClause(oTree(1))
    FormatDeleteStatement = FormatDeleteStatement & FormatFromClause(oTree(2), 0)
    If oTree(3).Index = 1 Then
        FormatDeleteStatement = FormatDeleteStatement & FormatWhereClause(oTree(3)(1), 0)
    End If
End Function

Private Function FormatDeleteClause(oTree As ParseTree) As String
    FormatDeleteClause = ChangeCase(oTree.Text, KeywordStyle) & msNewLine
End Function

Private Function FormatIntoClause(oTree As ParseTree, ByVal lIndentationLevel As Long)
    Dim sTableName As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
        
    FormatIntoClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    sTableName = ChangeCase(oTree(2).Text, TableStyle)
    If SpecialIdentifier(sTableName) Then
        FormatIntoClause = FormatIntoClause & sIndentation1 & "[" & sTableName & "]"
    Else
        FormatIntoClause = FormatIntoClause & sIndentation1 & sTableName
    End If
    FormatIntoClause = FormatIntoClause & msNewLine
End Function

Private Function FormatWithClause(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oSelect As ParseTree
    Dim sFields As String
    Dim sFormatFieldAlias As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim vStatements As Variant
    Dim sThisSelect As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatWithClause = sIndentation & ChangeCase(oTree(1)(1).Text, KeywordStyle) & " "
    vStatements = Array()
    For Each oSelect In oTree(1)(2).SubTree
        sThisSelect = FormatTable(oSelect(1)) & " " & ChangeCase(oSelect(2).Text, KeywordStyle) & " (" & msNewLine
        sThisSelect = sThisSelect & FormatSelectStatement(oSelect(4), lIndentationLevel + 1, False)
        sThisSelect = sThisSelect & sIndentation & ")"
        ReDim Preserve vStatements(UBound(vStatements) + 1)
        vStatements(UBound(vStatements)) = sThisSelect
    Next
    FormatWithClause = FormatWithClause & Join(vStatements, "," & msNewLine) & msNewLine
End Function

Private Function FormatSelectClause(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oField As ParseTree
    Dim sFields As String
    Dim sFormatFieldAlias As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatSelectClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & IIf(oTree(2).Index = 1, " " & ChangeCase(oTree(2).Text, KeywordStyle) & " ", "") & msNewLine

    For Each oField In oTree(3)(1).SubTree
        sFormatFieldAlias = FormatFieldAlias(oField)
        sFields = sFields & "," & msNewLine & sIndentation1 & sFormatFieldAlias
    Next
    sFields = Mid$(sFields, Len("," + msNewLine) + 1)
    FormatSelectClause = FormatSelectClause & sFields & msNewLine
End Function

Private Function FormatAnalyticFunction(oTree As ParseTree) As String
    Dim bIndentedTemp As Boolean
    Dim sOrderByClause As String
    Dim sTempNewLine As String
    
    sTempNewLine = msNewLine
    bIndentedTemp = mbIndented
    
    msNewLine = ""
    mbIndented = False
    
    FormatAnalyticFunction = FormatFunction(oTree(1))
    FormatAnalyticFunction = FormatAnalyticFunction & " " & ChangeCase(oTree(2).Text, KeywordStyle) & " ("
    If oTree(3).Index = 1 Then
        FormatAnalyticFunction = FormatAnalyticFunction & ChangeCase(oTree(3)(1)(1).Text, KeywordStyle) & " " & ChangeCase(oTree(3)(1)(2).Text, KeywordStyle) & " "
        FormatAnalyticFunction = FormatAnalyticFunction & FormatLevelX(1, oTree(3)(1)(3)) & " "
    End If
    If oTree(4).Index = 1 Then
        FormatAnalyticFunction = FormatAnalyticFunction & FormatOrderByClause(oTree(4)(1), 0)
    End If

    FormatAnalyticFunction = FormatAnalyticFunction & ")"
    
    msNewLine = sTempNewLine
    mbIndented = bIndentedTemp
End Function

Private Function FormatFieldAlias(oTree As ParseTree) As String
    Select Case oTree(1).Index
        Case 1
            FormatFieldAlias = FormatAnalyticFunction(oTree(1)(1))
        Case 2
            FormatFieldAlias = FormatLevelX(1, oTree(1)(1))
    End Select
    
    If oTree(2).Index = 1 Then
        If oTree(2)(1)(1).Index = 1 Then
            FormatFieldAlias = FormatFieldAlias & ChangeAS(oTree(2)(1)(1).Text, ASSelectStyle, KeywordStyle) & FormatField(oTree(2)(1)(8))
        Else
            FormatFieldAlias = FormatFieldAlias & ChangeAS("", ASSelectStyle, KeywordStyle) & FormatField(oTree(2)(1)(8))
        End If
    End If
End Function

Private Function FormatTableAlias(oTree As ParseTree) As String
    FormatTableAlias = FormatTable(oTree(1))

    If oTree(2).Index = 1 Then
        FormatTableAlias = FormatTableAlias & ChangeAS(oTree(2)(1)(1).Text, ASFromStyle, KeywordStyle) & oTree(2)(1)(13).Text
    End If
End Function

Private Function FormatField(oTree As ParseTree) As String
    If SpecialIdentifier(oTree.Text) Then
        Select Case Style
            Case 0
                FormatField = "[" & ChangeCase(oTree.Text, FieldStyle) & "]"
            Case 1
                FormatField = "`" & ChangeCase(oTree.Text, FieldStyle) & "`"
            Case 2
                FormatField = """" & ChangeCase(oTree.Text, FieldStyle) & """"
        End Select
    Else
        FormatField = ChangeCase(oTree.Text, FieldStyle)
    End If
End Function

Private Function FormatFromClause(oTree As ParseTree, ByVal lIndendationDepth As Long) As String
    Dim oTable As ParseTree
    Dim sTables As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndendationDepth)
    sIndentation1 = Indentation(lIndendationDepth + 1)
    
    FormatFromClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    
    For Each oTable In oTree(2)(1).SubTree
        sTables = sTables & "," & msNewLine & FormatJoinExpression(oTable, lIndendationDepth + 1)
    Next
    sTables = Mid$(sTables, 1 + Len("," & msNewLine))
    FormatFromClause = FormatFromClause & sTables & msNewLine
End Function

Private Function FormatJoinExpression(oTree As ParseTree, ByVal lIndentationDepth As Long, Optional ByVal bBrackets As Boolean) As String
    Dim iJoinLevel2Type As Integer
    Dim sIndentation As String
    
    sIndentation = Indentation(lIndentationDepth)
    
    iJoinLevel2Type = oTree(1).Index
    Select Case iJoinLevel2Type
        Case 1
            FormatJoinExpression = StripNewLine(FormatUnionStatement(oTree(1)(1)(1)(1)(1), lIndentationDepth, True))
            If oTree(1)(1)(3).Index = 1 Then
                FormatJoinExpression = FormatJoinExpression & ChangeCase(ChangeAS(oTree(1)(1)(3)(1)(1).Text, ASFromStyle, KeywordStyle), KeywordStyle) & oTree(1)(1)(3)(1)(12).Text
            End If
        Case 2
            FormatJoinExpression = FormatJoinExpression(oTree(1)(1)(1), lIndentationDepth + 1, True)
        Case 3
            FormatJoinExpression = sIndentation & FormatTableAlias(oTree(1)(1))
            If Not colTables.Exists(LCase$(oTree(1)(1)(1).Text)) Then
                colTables.Add LCase$(oTree(1)(1)(1).Text), oTree(1)(1)(1).Text
            End If
            If Not colTableAliases.Exists(LCase$(oTree(1)(1)(1).Text)) Then
                If oTree(1)(1)(2).Index = 1 Then
                    colTableAliases.Add LCase$(oTree(1)(1)(1).Text), oTree(1)(1)(2)(1)(13).Text
                    If Not colTableUnaliases.Exists(LCase$(oTree(1)(1)(2)(1)(13).Text)) Then
                        colTableUnaliases.Add LCase$(oTree(1)(1)(2)(1)(13).Text), oTree(1)(1)(1).Text
                    End If
                End If
            End If
    End Select

    If oTree(2).Index = 1 Then
        FormatJoinExpression = FormatJoinExpression & IIf(Indented = 1, vbCrLf, "") & sIndentation & ChangeCase(oTree(2)(1)(1).Text, KeywordStyle) & ChangeCase(oTree(2)(1)(2).Text, KeywordStyle) & IIf(Indented = 1, vbCrLf, "")
        iJoinLevel2Type = oTree(2)(1)(3).Index
        Select Case iJoinLevel2Type
            Case 1
                FormatJoinExpression = FormatJoinExpression & StripNewLine(FormatUnionStatement(oTree(2)(1)(3)(1)(1)(1)(1), lIndentationDepth, True))
                If oTree(2)(1)(3)(1)(3).Index = 1 Then
                    FormatJoinExpression = FormatJoinExpression & ChangeCase(ChangeAS(oTree(2)(1)(3)(1)(3)(1)(1).Text, ASFromStyle, KeywordStyle), KeywordStyle) & oTree(2)(1)(3)(1)(3)(1)(12).Text
                End If
            Case 2
                FormatJoinExpression = FormatJoinExpression & FormatJoinExpression(oTree(2)(1)(3)(1)(1), lIndentationDepth + 1, True)
            Case 3
                FormatJoinExpression = FormatJoinExpression & sIndentation & FormatTableAlias(oTree(2)(1)(3)(1))
                If Not colTables.Exists(LCase$(oTree(2)(1)(3)(1)(1).Text)) Then
                    colTables.Add LCase$(oTree(2)(1)(3)(1)(1).Text), oTree(2)(1)(3)(1)(1).Text
                End If
                If Not colTableAliases.Exists(LCase$(oTree(2)(1)(3)(1)(1).Text)) Then
                    If oTree(2)(1)(3)(1)(2).Index = 1 Then
                        colTableAliases.Add LCase$(oTree(2)(1)(3)(1)(1).Text), oTree(2)(1)(3)(1)(2)(1)(13).Text
                        colTableUnaliases.Add LCase$(oTree(2)(1)(3)(1)(2)(1)(13).Text), oTree(2)(1)(3)(1)(1).Text
                    End If
                End If
        End Select
        FormatJoinExpression = FormatJoinExpression & IIf(Indented = 1, vbCrLf, "") & sIndentation & ChangeCase(oTree(2)(1)(4).Text, KeywordStyle) & " " & FormatLevelX(1, oTree(2)(1)(5))
    End If
    
    If bBrackets Then
        FormatJoinExpression = sIndentation & "(" & StripIndentation(FormatJoinExpression, lIndentationDepth) & ")"
    End If
End Function

Private Function FormatJoinBrackets(ByVal sIndentLevel, ByVal bBrackets, ByVal sExpression As String) As String
    'FormatJoinBrackets = sIndentLevel & IIf(bBrackets, "(", "") & Mid$(sExpression, Len(sIndentLevel) + 1) & IIf(bBrackets, ")", "")
    FormatJoinBrackets = sIndentLevel & IIf(bBrackets, "(", "") & sExpression & IIf(bBrackets, ")", "")
End Function

Private Function FormatTable(oTree As ParseTree) As String
    If SpecialIdentifier(oTree.Text) Then
        If Style = 0 Then
            FormatTable = "[" & ChangeCase(oTree.Text, TableStyle) & "]"
        Else
            FormatTable = "`" & ChangeCase(oTree.Text, TableStyle) & "`"
        End If
    Else
        FormatTable = ChangeCase(oTree.Text, TableStyle)
    End If
End Function

Private Sub InitialiseSpecialCharacters()
    Dim iIndex As Integer
    
    For iIndex = 32 To 35
        sSpecialCharacters = sSpecialCharacters & Chr$(iIndex)
    Next
    
    For iIndex = 37 To 47
        sSpecialCharacters = sSpecialCharacters & Chr$(iIndex)
    Next
    
    For iIndex = 58 To 63
        sSpecialCharacters = sSpecialCharacters & Chr$(iIndex)
    Next
        
    For iIndex = 91 To 94
        sSpecialCharacters = sSpecialCharacters & Chr$(iIndex)
    Next
        
    For iIndex = 123 To 255
        sSpecialCharacters = sSpecialCharacters & Chr$(iIndex)
    Next
End Sub

Private Function SpecialIdentifier(sIdentifier As String) As Boolean
    Dim iPos As Integer

    If sIdentifier = "*" Then
        Exit Function
    End If
    
    For iPos = 1 To Len(sIdentifier)
        If InStr(sSpecialCharacters, Mid$(sIdentifier, iPos, 1)) > 0 Then
            SpecialIdentifier = True
        End If
    Next
End Function

Private Function FormatWhereClause(oTree As ParseTree, ByVal lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
    
    FormatWhereClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    If WHEREExpression = 0 Then
        FormatWhereClause = FormatWhereClause & sIndentation1 & FormatLevelX(1, oTree(2), 0, 0, lIndentationDepth + 1, 2) & msNewLine
    Else
        FormatWhereClause = FormatWhereClause & sIndentation1 & FormatLevelX(1, oTree(2), 0, 0, lIndentationDepth + 1) & msNewLine
    End If
End Function

Private Function FormatFieldQualifiedBrackets(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1
            FormatFieldQualifiedBrackets = FormatQualifiedField(oTree(1))
        Case 2
            FormatFieldQualifiedBrackets = FormatFieldQualifiedBrackets(oTree(1)(1))
    End Select
End Function

Private Function FormatLevelX(ByVal lLevel As Long, oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long, Optional lIndentationDepth As Long, Optional lIndentationTargetDepth As Long) As String
    Dim oTerm As ParseTree
    Dim iOperands As Integer
    Dim sIndentation As String
    
    sIndentation = Indentation(lIndentationDepth)
    
    iOperands = (oTree.SubTree.Count + 1) \ 2
    If iOperands > 1 Then
        lCallingDepth = lLevel
    End If
    If lLevel <= 5 Then
        For Each oTerm In oTree.SubTree
            Select Case oTerm.Description
                Case "AND" ' operator
                    If lIndentationDepth > 0 And lLevel = lIndentationTargetDepth Then
                        FormatLevelX = FormatLevelX & msNewLine & sIndentation & ChangeCase(oTerm.Text, OperatorStyle) & " "
                    Else
                        FormatLevelX = FormatLevelX & " " & ChangeCase(oTerm.Text, OperatorStyle) & " "
                    End If
                    If lLevel <= lDepth Then
                        lDepth = lLevel
                    End If
                Case Else ' operand
                    FormatLevelX = FormatLevelX & FormatLevelX(lLevel + 1, oTerm, lDepth, lCallingDepth, lIndentationDepth, lIndentationTargetDepth)
            End Select
        Next
    Else
        For Each oTerm In oTree.SubTree
            If oTerm.Name = "level6" Then
                FormatLevelX = FormatLevelX & FormatLevel6(oTerm, lDepth, lCallingDepth)
            Else
                FormatLevelX = FormatLevelX & " " & UCase$(oTerm.Text) & " "
            End If
            If lLevel <= lDepth Then
                lDepth = lLevel
            End If
        Next
    End If
End Function

Private Function FormatLevel6(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    Dim sExpression As String
    Dim iTestDepth As Integer
    
    FormatLevel6 = FormatLevel7(oTree(1), lDepth, lCallingDepth)
    
    If oTree(2).Index = 1 Then
        Select Case oTree(2)(1).Index
            Case 1 ' is
                FormatLevel6 = FormatLevel6 & FormatIs(oTree(2)(1)(1), lDepth, lCallingDepth)
            Case 2 ' in
                FormatLevel6 = FormatLevel6 & FormatIn(oTree(2)(1)(1), lDepth, lCallingDepth)
            Case 3 ' like
                FormatLevel6 = FormatLevel6 & FormatLike(oTree(2)(1)(1), lDepth, lCallingDepth)
            Case 4 ' between
                FormatLevel6 = FormatLevel6 & FormatBetween(oTree(2)(1)(1), lDepth, lCallingDepth)
            Case 5 ' level8
                FormatLevel6 = FormatLevel6 & FormatLevel7(oTree(2)(1)(1), lDepth, lCallingDepth)
        End Select
    End If
End Function

Private Function FormatLevel7(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    Dim sExpression As String
    Dim lTestDepth As Long
    
    Select Case oTree.Index
        Case 1 ' date
            FormatLevel7 = oTree(1).Text
        Case 2 'string
            FormatLevel7 = FormatString(oTree(1))
        Case 3 ' unary
            FormatLevel7 = FormatUnaryValue(oTree(1))
        Case 4 ' number
            FormatLevel7 = oTree(1).Text
        Case 5 ' function
            FormatLevel7 = FormatFunction(oTree(1))
        Case 6 ' qualified field
            FormatLevel7 = FormatQualifiedField(oTree(1))
        Case 7 ' simple sub query
            FormatLevel7 = FormatUnionStatement(oTree(1)(1)(1)(1), 0, True)
        Case 8 ' brackets
            lTestDepth = 7
            sExpression = FormatLevelX(1, oTree(1)(1), lTestDepth)
            If lTestDepth <= lCallingDepth Then
                sExpression = "(" & sExpression & ")"
            End If
            FormatLevel7 = sExpression
    End Select
End Function

Private Function FormatIs(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    FormatIs = " IS " & IIf(oTree(1).Index = 1, "NOT ", "") & "NULL"
End Function

Private Function FormatBetween(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    FormatBetween = IIf(oTree(1).Index = 1, " NOT", "") & " BETWEEN " & FormatLevelX(5, oTree(2)) & " AND " & FormatLevelX(5, oTree(3))
End Function

Private Function FormatIn(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    Dim oElement As ParseTree
    Dim sStrings As String
    Dim lIndentedTemp As Long
    
    Dim bIndented As Boolean
    Dim sNewLine As String
    
    FormatIn = IIf(oTree(1).Index = 1, " NOT", "") & " IN "
    
    Select Case oTree(2).Index
        Case 1 'LIST
            For Each oElement In oTree(2)(1)(1).SubTree
                Select Case oElement.Index
                    Case 1
                        sStrings = sStrings & ", " & FormatString(oElement)
                    Case 2
                        sStrings = sStrings & ", " & oElement.Text
                End Select
            Next
            FormatIn = FormatIn & "(" & Mid$(sStrings, 3) & ")"
        Case 2 'SELECT
            bIndented = mbIndented
            sNewLine = msNewLine
            mbIndented = IndentedSubQuery = 1
            msNewLine = IIf(mbIndented, vbCrLf, "")
            FormatIn = FormatIn & FormatUnionStatement(oTree(2)(1)(1)(1)(1), 0, True)
            mbIndented = bIndented
            msNewLine = sNewLine
    End Select
End Function

Private Function FormatLike(oTree As ParseTree, Optional ByRef lDepth As Long, Optional ByVal lCallingDepth As Long) As String
    FormatLike = IIf(oTree(1).Index = 1, " NOT", "") & " LIKE " & oTree(2).Text
End Function


Private Function FormatString(oTree As ParseTree) As String
    FormatString = ChangeQuotes(oTree.Text, QuoteStyle)
End Function

Private Function FormatUnaryValue(oTree As ParseTree) As String
    FormatUnaryValue = oTree(1).Text & IIf(UCase$(oTree(1).Text) = "NOT", " ", "") & FormatLevelX(1, oTree(2))
End Function

Private Function FormatQualifiedField(oTree As ParseTree) As String
    Dim sTableName As String
    
    Dim lFieldCount As Long
    Dim lFieldIndex As Long
    Dim oField As ParseTree
    Dim sTable As String
    
    lFieldCount = oTree(1)(1).Index
    lFieldIndex = 1
    For Each oField In oTree(1)(1).SubTree
        If lFieldIndex = lFieldCount And oField(2).Text = "." Then 'Table
            sTable = oField(1).Text

            Select Case TableAliasing
                Case 0 ' Aliased
                    If colTableAliases.Exists(LCase$(sTable)) Then
                        FormatQualifiedField = FormatQualifiedField & colTableAliases.Item(LCase$(sTable)) & "."
                    Else
                        FormatQualifiedField = FormatQualifiedField & FormatTable(oField(1)) & "."
                    End If
                Case 1 ' Unalias
                    If colTableUnaliases.Exists(LCase$(sTable)) Then
                        FormatQualifiedField = FormatQualifiedField & colTableUnaliases.Item(LCase$(sTable)) & "."
                    Else
                        FormatQualifiedField = FormatQualifiedField & FormatTable(oField(1)) & "."
                    End If
                Case 2 ' No aliasing
                    FormatQualifiedField = FormatQualifiedField & ""
                Case 3 'Leave
                    FormatQualifiedField = FormatQualifiedField & FormatTable(oField(1)) & "."
            End Select
        Else
            FormatQualifiedField = FormatQualifiedField & FormatTable(oField(1)) & oField(2).Text
        End If
        
        lFieldIndex = lFieldIndex + 1
    Next
    
    FormatQualifiedField = FormatQualifiedField & FormatField(oTree(1)(2))
    
    If oTree(2).Index = 1 Then
        FormatQualifiedField = FormatQualifiedField & " (+)"
    End If
End Function

Private Function FormatFunction(oTree As ParseTree) As String
    Dim oParameter As ParseTree
    Dim sParameterList As String
    
    For Each oParameter In oTree(2).SubTree
        sParameterList = sParameterList & ", " & FormatLevelX(1, oParameter)
    Next
    FormatFunction = ChangeCase(oTree(1).Text, FunctionStyle) & "(" & Mid$(sParameterList, 3) & ")"
End Function

Private Function FormatGroupByClause(oTree As ParseTree, lIndentationDepth As Long) As String
    Dim oField As ParseTree
    Dim sFields As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
    
    FormatGroupByClause = sIndentation & ChangeCase(oTree(1).Text & " " & oTree(2).Text, KeywordStyle) & msNewLine
    
    For Each oField In oTree(3).SubTree
        sFields = sFields & "," & msNewLine & sIndentation1 & FormatLevelX(1, oField(1))
        If oField(2).Index = 1 Then
            sFields = sFields & " " & ChangeCase(oField(2).Text, KeywordStyle)
        End If
        If oField(3).Index = 1 Then
            sFields = sFields & " " & ChangeCase(oField(3)(1)(1).Text, KeywordStyle) & " " & ChangeCase(oField(3)(1)(2).Text, KeywordStyle)
        End If
    Next
    sFields = Mid$(sFields, 1 + Len("," & msNewLine))
    FormatGroupByClause = FormatGroupByClause & sFields & msNewLine
End Function

Private Function FormatHavingClause(oTree As ParseTree, lIndentationDepth As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
        
    FormatHavingClause = sIndentation & ChangeCase(oTree(1).Text, KeywordStyle) & msNewLine
    FormatHavingClause = FormatHavingClause & sIndentation1 & FormatLevelX(1, oTree(2), 0, 0, lIndentationDepth + 1, True) & msNewLine
End Function

Private Function FormatOrderByClause(oTree As ParseTree, lIndentationDepth As Long) As String
    Dim oField As ParseTree
    Dim sFields As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationDepth)
    sIndentation1 = Indentation(lIndentationDepth + 1)
        
    FormatOrderByClause = sIndentation & ChangeCase(oTree(1).Text & " " & oTree(2).Text, KeywordStyle) & msNewLine
    
    For Each oField In oTree(3).SubTree
        sFields = sFields & "," & msNewLine & sIndentation1 & FormatLevelX(1, oField(1))
        If oField(2).Index = 1 Then
            sFields = sFields & " " & ChangeCase(oField(2)(1).Text, KeywordStyle)
        End If
        If oField(3).Index = 1 Then
            sFields = sFields & " " & ChangeCase(oField(3)(1)(1).Text, KeywordStyle) & " " & ChangeCase(oField(3)(1)(2).Text, KeywordStyle)
        End If
    Next
    sFields = Mid$(sFields, 1 + Len("," & msNewLine))
    FormatOrderByClause = FormatOrderByClause & sFields & msNewLine
End Function

Private Function ChangeCase(ByVal sText As String, iIndex As Integer) As String
    Select Case iIndex
        Case 0 ' UPPER
            ChangeCase = UCase$(sText)
        Case 1 ' LOWER
            ChangeCase = LCase$(sText)
        Case 2 ' PROPER
            ChangeCase = StrConv(sText, vbProperCase)
        Case 3 ' LEAVE
            ChangeCase = sText
    End Select
End Function

Private Function ChangeQuotes(ByVal sText As String, iIndex As Integer) As String
    Select Case iIndex
        Case 0 ' SINGLE
            ChangeQuotes = "'" & Mid$(sText, 2, Len(sText) - 2) & "'"
        Case 1 ' DOUBLE
            ChangeQuotes = """" & Mid$(sText, 2, Len(sText) - 2) & """"
        Case 2 ' LEAVE
            ChangeQuotes = sText
    End Select
End Function

Private Function ChangeAS(ByVal sText As String, iIndex As Integer, iIndex2 As Integer) As String
    Select Case iIndex
        Case 0 ' BLANK
            ChangeAS = " "
        Case 1 ' AS
            ChangeAS = ChangeCase(" AS ", iIndex2)
        Case 2 ' LEAVE
            If sText <> "" Then
                ChangeAS = " " & ChangeCase(sText, iIndex2) & " "
            Else
                ChangeAS = " "
            End If
    End Select
End Function

Private Function Indentation(ByVal lDepth As Long) As String
    Indentation = IIf(mbIndented, String$(lDepth * 4, " "), " ")
End Function

Private Function StripIndentation(sString As String, ByVal lDepth As Long) As String
    StripIndentation = Mid$(sString, 1 + Len(Indentation(lDepth)))
End Function

Private Function StripNewLine(sString As String) As String
    StripNewLine = Left$(sString, Len(sString) - Len(msNewLine))
End Function
