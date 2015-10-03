Attribute VB_Name = "KwikSyntax"
Option Explicit

Public oStripQuotes As IParseObject
Public LineParse As IParseObject
Public EndOfStatement As IParseObject

Private Positions As Collection
Private Colours As Collection
Private CanonicalText As String

Private keyword_col As Long
Private variable_col As Long
Private bracket_col As Long
Private variableassign_col As Long
Private auxkeyword_col As Long

Public Sub InitialiseParser()
    Dim Definition As String
    
    Definition = "ws :=AND ?{REPEAT IN 1 TO 32 UNTIL 13+ 10 MIN 0}, {REPEAT IN 1 TO 32 MIN 0};" & _
                     "ws2 :=[AND  (REPEAT IN 1 TO 32 UNTIL OR EOS, 13+10 MIN 0), (OR EOS, 13+10)];" & _
                     "variable := AND (IN CASE 'A' TO 'Z', '_'), REPEAT (IN CASE 'A' TO 'Z', '0' TO '9', '_') MIN 0;" & _
                     "nonalpha := IN CASE 'A' TO 'Z', '0' TO '9','_';" & _
                     "number := REPEAT IN '0' TO '9' MAX 3;" & _
                     "string := AND 39, (REPEAT OR (AND 39, 39), (IN 32 TO 255) UNTIL AND 39, NOT 39), 39;" & _
                     "st_literal := AND OPTIONAL (AND (OR CASE 'CASE', '^'), ws), (LIST (OR number, string), (AND '+', ws));" & _
                     "char := OR number, (AND 39, (IN 32 TO 255, NOT 39), 39);" & _
                     "range := AND char, ws, (OR CASE 'TO', '-'), ws, char;" & _
                     "st_in := AND (OR CASE 'IN', '>'), ws, (LIST (AND (OPTIONAL (OR CASE 'NOT', '!')), ws, (OPTIONAL (OR CASE 'CASE', '^')), ws, (OR range, number, string)), (AND ',', ws));" & _
                     "st_not := AND (OR (AND (CASE 'NOT'), NOT nonalpha), '!'), ws, expression;" & _
                     "st_and := AND (OR (AND (CASE 'AND'), NOT nonalpha), '&'), ws, (LIST expression, (AND ws,',', ws) MIN 1);" & _
                     "st_or := AND (OR (AND (CASE 'OR'), NOT nonalpha), '|'), ws, (LIST expression, (AND ws,',', ws) MIN 1);"
    Definition = Definition & _
                     "st_repeat := AND (OR (AND (CASE 'REPEAT'), NOT nonalpha), '#'), ws, expression, (OPTIONAL AND ws, (OR CASE 'UNTIL', ':'), ws, expression), (OPTIONAL AND ws, (OR CASE 'MIN', '-'), ws, number), (OPTIONAL AND ws, (OR CASE 'MAX', '+'), ws, number);" & _
                     "st_list := AND (OR (AND (CASE 'LIST'), NOT nonalpha), '@'), ws, expression, ws, ',', ws, expression, (OPTIONAL AND ws, (OR CASE 'MIN', '-'), ws, number), (OPTIONAL AND ws, (OR CASE 'MAX', '+'), ws, number);" & _
                     "st_optional := AND (OR (AND (CASE 'OPTIONAL'), NOT nonalpha), '?'), ws, expression;" & _
                     "parameter_string := {REPEAT IN 'A' TO 'Z', 'a' TO 'z', '_'};" & _
                     "parameter_list := LIST parameter_string, [','];" & _
                     "st_external := AND (OR (AND (CASE 'EXTERNAL'), NOT nonalpha), '='), ws, variable, (OPTIONAL AND ['('], parameter_list, [')']);" & _
                     "expression_b := AND '(', ws, expression_sub, ws, ')';" & _
                     "expression_sb := AND '[', ws, expression_sub, ws, ']';" & _
                     "expression_cb := AND '{', ws, expression_sub, ws, '}';" & _
                     "expression := OR expression_b, expression_sb, expression_cb, expression_sub;" & _
                     "expression_sub := OR st_literal, st_and, st_or, st_repeat, st_list, st_in, st_optional, st_not, (OR CASE 'EOS', '||'), (OR CASE 'PASS', '*'), (OR CASE 'FAIL', '~'), st_external, variable;" & _
                     "statement := AND REPEAT 32 MIN 0, variable, ws, ':=', ws, expression, ws, ';', ws2;" & _
                     "endofstatement := AND [REPEAT IN 32 TO 255, NOT ';'], ';', ws, EOS;"
    Definition = Definition & "strip_quotes := {AND [?'""'], REPEAT IN 32 TO 255 UNTIL (OR EOS, '""')};"
                     
    If Not SetNewDefinition(Definition) Then
        Debug.Print ErrorString
        End
    End If
    
    Set oStripQuotes = ParserObjects("strip_quotes")
    Set LineParse = ParserObjects("statement")
    Set EndOfStatement = ParserObjects("endofstatement")
    
    keyword_col = RGB(0, 0, 160) ' navy
    variable_col = RGB(96, 96, 96) ' dark grey
    bracket_col = RGB(128, 0, 0) ' dark red
    variableassign_col = RGB(0, 128, 0) ' dark green
    auxkeyword_col = RGB(128, 0, 128) ' dark magenta
End Sub

Public Function ParseLine(linetext As String, Optional iOutputPos, Optional bParsed As Boolean) As Variant
    Dim oParseTree As New ParseTree
    Dim vParseExpression As Variant
    Dim vPosition As Variant
    Dim lindex As Long
    Dim member As ParseTree
    Dim bSpacer As Boolean
    
    ParserText = linetext
    
    Set Positions = New Collection
    Set Colours = New Collection
    CanonicalText = ""
    
    If Not LineParse.Parse(oParseTree) Then
        If linetext <> "" Then
            AddText linetext, vbRed
        End If
        ParseLine = Array(CanonicalText, Positions, Colours)
        iOutputPos = Len(linetext) + 1
        bParsed = False
        Exit Function
    End If
    bParsed = True
    iOutputPos = ParserTextPosition
    
    ' Variable
    AddText oParseTree(1).Text & oParseTree(2).Text, variableassign_col
    
    ' ws
    bSpacer = WS(oParseTree(3))
    
    ' Colon equals
    AddText Spacer(bSpacer) & ":= ", vbBlack
    
    ' ws
    WS oParseTree(5)
    
    Select Case oParseTree(6).index
        Case 1, 2, 3
            AddText oParseTree(6)(1)(1).Text, bracket_col
            Set member = oParseTree(6)(1)(3)
            ParseExpression member
            AddText oParseTree(6)(1)(5).Text, bracket_col
        Case Else
            Set member = oParseTree(6)(1)
            ParseExpression member
    End Select

    ' Semi colon
    AddText ";", vbBlack
    
    ParseLine = Array(CanonicalText, Positions, Colours)
End Function

Private Function WS(oWS As ParseTree) As Boolean
    If oWS(1).index = 1 Then
        AddText oWS(2).Text, vbBlack
    Else
        WS = True
    End If
End Function

Private Function Spacer(bYes As Boolean)
    If bYes Then
        Spacer = " "
    End If
End Function

Private Function ParseExpression(oResult As ParseTree)
    Dim member As ParseTree
    Dim lindex As Long
    
    Select Case oResult.index
        Case 0 ' bracketed expression
            AddText oResult(1).Text, bracket_col
            WS oResult(2)
            Set member = oResult(3)
            ParseExpression member
            WS oResult(4)
            AddText oResult(5).Text, bracket_col
            
        Case 1 ' literal
            Set member = oResult(1)
            ParseLiteralSubExpression member
                                    
        Case 2, 3 ' and, or
            If oResult(1)(1).index = 1 Then
                AddText UCase(oResult(1)(1).Text) & " ", keyword_col
            Else
                AddText oResult(1)(1).Text, keyword_col
            End If
            
            WS oResult(1)(2)
            
            For lindex = 1 To oResult(1)(3).SubTree.Count
                Set member = oResult(1)(3)(lindex)
                If member.Name = "expression" Then
                    ParseExpression member(1)
                Else
                    WS member(1)
                    AddText ",", vbBlack
                    AddText Spacer(WS(member(3))), vbBlack
                End If
            Next
            
        Case 4 'repeat
            If oResult(1)(1).index = 1 Then
                AddText "REPEAT", keyword_col
            Else
                AddText "#", keyword_col
            End If
            
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            
            Set member = oResult(1)(3)(1)
            ParseExpression member
            
            If oResult(1)(4).index = 1 Then
                AddText Spacer(WS(oResult(1)(4)(1)(1))), vbBlack
                If oResult(1)(4)(1)(2).index = 1 Then
                    AddText "UNTIL", keyword_col
                Else
                    AddText ":", keyword_col
                End If
                AddText Spacer(WS(oResult(1)(4)(1)(3))), vbBlack
                Set member = oResult(1)(4)(1)(4)(1)
                ParseExpression member
            End If
            
            If oResult(1)(5).index = 1 Then
                AddText Spacer(WS(oResult(1)(5)(1)(1))), vbBlack
                If oResult(1)(5)(1)(2).index = 1 Then
                    AddText "MIN", keyword_col
                Else
                    AddText "-", keyword_col
                End If
                AddText Spacer(WS(oResult(1)(5)(1)(3))), vbBlack
                AddText oResult(1)(5)(1)(4).Text, vbBlack
            End If
            
            If oResult(1)(6).index = 1 Then
                AddText Spacer(WS(oResult(1)(6)(1)(1))), vbBlack
                If oResult(1)(6)(1)(2).index = 1 Then
                    AddText "MAX", keyword_col
                Else
                    AddText "+", keyword_col
                End If
                AddText Spacer(WS(oResult(1)(6)(1)(3))), vbBlack
                AddText oResult(1)(6)(1)(4).Text, vbBlack
            End If
            
        Case 5 'list
            If oResult(1)(1).index = 1 Then
                AddText "LIST", keyword_col
            Else
                AddText "@", keyword_col
            End If
        
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            
            Set member = oResult(1)(3)(1)
            ParseExpression member
            WS oResult(1)(4)
            AddText ",", vbBlack
            AddText Spacer(WS(oResult(1)(6))), vbBlack
            Set member = oResult(1)(7)(1)
            ParseExpression member
            
            If oResult(1)(8).index = 1 Then
                AddText Spacer(WS(oResult(1)(8)(1)(1))), vbBlack
                If oResult(1)(8)(1)(2).index = 1 Then
                    AddText "MIN", keyword_col
                Else
                    AddText "-", keyword_col
                End If
                AddText Spacer(WS(oResult(1)(8)(1)(3))), vbBlack
                AddText oResult(1)(8)(1)(4).Text, vbBlack
            End If
            
            If oResult(1)(9).index = 1 Then
                AddText Spacer(WS(oResult(1)(9)(1)(1))), vbBlack
                If oResult(1)(9)(1)(2).index = 1 Then
                    AddText "MAX", keyword_col
                Else
                    AddText "+", keyword_col
                End If
                AddText Spacer(WS(oResult(1)(9)(1)(3))), vbBlack
                AddText oResult(1)(9)(1)(4).Text, vbBlack
            End If
            
        Case 6 'in
            If oResult(1)(1).index = 1 Then
                AddText "IN", keyword_col
            Else
                AddText ">", keyword_col
            End If
            
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            Set member = oResult(1)(3)
            ParseInSubExpression member
            
        Case 7 ' optional
            If oResult(1)(1).index = 1 Then
                AddText "OPTIONAL", keyword_col
            Else
                AddText "?", keyword_col
            End If
            
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            Set member = oResult(1)(3)(1)
            ParseExpression member
            
        Case 8 ' not
            If oResult(1)(1).index = 1 Then
                AddText "NOT", keyword_col
            Else
                AddText "!", keyword_col
            End If
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            Set member = oResult(1)(3)(1)
            ParseExpression member
            
        Case 9 ' EOS
            If oResult(1).index = 1 Then
                AddText "EOS", auxkeyword_col
            Else
                AddText "||", auxkeyword_col
            End If
        
        Case 10 ' PASS
            If oResult(1).index = 1 Then
                AddText "PASS", auxkeyword_col
            Else
                AddText "*", auxkeyword_col
            End If
            
        Case 11 'FAIL
            If oResult(1).index = 1 Then
                AddText "FAIL", auxkeyword_col
            Else
                AddText "~", auxkeyword_col
            End If
            
        Case 12 ' External
            If oResult(1)(1).index = 1 Then
                AddText "EXTERNAL", keyword_col
            Else
                AddText "=", keyword_col
            End If
            AddText Spacer(WS(oResult(1)(2))), vbBlack
            
            AddText oResult(1)(3).Text, vbBlack
            
            Dim sParameterText As String
            
            Select Case oResult(1)(4).index
                Case 0 ' No parameters
                Case 1
                    Dim vParameter As Variant
                    AddText "(", vbBlack
                    For Each vParameter In oResult(1)(4)(1)(1).SubTree
                        sParameterText = sParameterText & "," & vParameter.Text
                    Next
                    
                    AddText Mid$(sParameterText, 2), vbBlack
                    AddText ")", vbBlack
            End Select
            
        Case 13 'variable
            AddText oResult(1).Text, variable_col
    End Select
End Function

Private Function ParseInSubExpression(oResult As ParseTree)
    Dim oSub As ParseTree
    
    For Each oSub In oResult.SubTree
        If oSub(1).Text = "," Then
            AddText ",", vbBlack
            AddText Spacer(WS(oSub(2))), vbBlack
        Else
            If oSub(1).index = 1 Then
                If oSub(1)(1).index = 1 Then
                    AddText "NOT", auxkeyword_col
                Else
                    AddText "!", auxkeyword_col
                End If
                AddText Spacer(WS(oSub(2))), vbBlack
            End If
            
            If oSub(3).index = 1 Then
                If oSub(3)(1).index = 1 Then
                    AddText "CASE", auxkeyword_col
                Else
                    AddText "^", auxkeyword_col
                End If
                AddText Spacer(WS(oSub(4))), vbBlack
            End If
            
            Select Case oSub(5).index
                Case 1 ' range
                    AddText oSub(5)(1)(1).Text, vbBlack
                    AddText Spacer(WS(oSub(5)(1)(2))), vbBlack
                    If oSub(5)(1)(3).index = 1 Then
                        AddText "TO", auxkeyword_col
                    Else
                        AddText "-", auxkeyword_col
                    End If
                    AddText Spacer(WS(oSub(5)(1)(4))), vbBlack
                    AddText oSub(5)(1)(5).Text, vbBlack
                Case 2, 3 ' number
                    AddText oSub(5)(1).Text, vbBlack
            End Select
        End If
    Next
End Function

Private Function ParseLiteralSubExpression(oResult As ParseTree)
    Dim oSub As ParseTree
    
    If oResult(1).index = 1 Then
        If oResult(1)(1)(1).index = 1 Then
            AddText "CASE ", auxkeyword_col
        Else
            AddText "^", auxkeyword_col
        End If
    End If
    
    For Each oSub In oResult(2).SubTree
        If oSub(1).Text = "+" Then
            AddText "+ ", auxkeyword_col
        Else
            AddText oSub.Text, vbBlack
        End If
    Next
End Function

Private Sub AddText(ByVal sAddString As String, ByVal lColour As Long)
    Dim lTextPos As Long
    
    lTextPos = Len(CanonicalText) + 1
    CanonicalText = CanonicalText & sAddString
    Positions.Add lTextPos
    Colours.Add lColour
End Sub

Public Function EndOfStatementText(ByVal sStatement) As Boolean
    Dim oResult As New ParseTree
    
    ParserText = sStatement
    EndOfStatementText = EndOfStatement.Parse(oResult)
End Function
