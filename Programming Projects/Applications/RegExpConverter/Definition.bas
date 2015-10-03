Attribute VB_Name = "Definition"
Option Explicit

Public oParser As IParseObject
Public AllSets As New Dictionary

Public Sub InitialiseParser()
    Dim oFS As New FileSystemObject
    
    Dim sDef As String

    sDef = oFS.OpenTextFile(App.Path & "\RegExpConverter.pdl").ReadAll
            
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParser = ParserObjects("expression")
End Sub

Public Function ConvertRegExp(oTree As ParseTree) As String
    ConvertRegExp = "/" & Expression(oTree) & "/"
End Function

Private Function Expression(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1, 2, 3, 4 ' brackets
            Expression = ExpressionSub(oTree(1)(2))
        Case 5 ' sub
            Expression = ExpressionSub(oTree(1))
    End Select
End Function

Private Function ExpressionSub(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1 ' literal
            ExpressionSub = StLiteral(oTree(1))
        Case 2 ' and
            ExpressionSub = StAnd(oTree(1))
        Case 3 ' perm
            ExpressionSub = StPerm(oTree(1))
        Case 4 ' or
            ExpressionSub = StOr(oTree(1))
        Case 5 ' repeat
            ExpressionSub = StRepeat(oTree(1))
        Case 6 ' list
            ExpressionSub = StList(oTree(1))
        Case 7 ' in
            ExpressionSub = StIn(oTree(1))
        Case 8 ' optional
            ExpressionSub = StOptional(oTree(1))
        Case 9 ' not
            ExpressionSub = StNot(oTree(1))
        Case 10 ' each
            ExpressionSub = StEach(oTree(1))
        Case 11 ' eos
            ExpressionSub = StEOS(oTree(1))
        Case 12 ' bos
            ExpressionSub = StBOS(oTree(1))
        Case 13 ' pass
            ExpressionSub = StPass(oTree(1))
        Case 14 ' fail
            ExpressionSub = StFail(oTree(1))
        Case 15 ' external
            ExpressionSub = StExternal(oTree(1))
        Case 16 ' variable
            ExpressionSub = StVariable(oTree(1))
    End Select
End Function

Private Function StLiteral(oTree As ParseTree) As String
    Dim oPart As ParseTree
    
    Select Case oTree(1).Index
        Case 1 ' Case
    End Select
    
    For Each oPart In oTree(2).SubTree
        Select Case oPart.Index
            Case 1 ' number
            Case 2 ' string
                StLiteral = StLiteral & EscapeString(oPart.Text)
        End Select
    Next
End Function

Private Function EscapeString(sString As String) As String
    EscapeString = Replace$(sString, "\", "\\\")
    EscapeString = Replace$(EscapeString, ".", "\.")
    EscapeString = Replace$(EscapeString, "^", "\^")
    EscapeString = Replace$(EscapeString, "$", "\$")
    EscapeString = Replace$(EscapeString, "/", "\/")
End Function

Private Function StAnd(oTree As ParseTree) As String
    Dim oPart As ParseTree
    
    Select Case oTree(1).Index
        Case 1 ' pass /fail
    End Select
        
    For Each oPart In oTree(2).SubTree
        StAnd = StAnd & Expression(oPart)
    Next
    StAnd = "(" & StAnd & ")"
End Function

Private Function StPerm(oTree As ParseTree) As String

End Function

Private Function StOr(oTree As ParseTree) As String
    Dim oPart As ParseTree
    Dim vParts As Variant
    
    vParts = Array()
    
    Select Case oTree(1).Index
        Case 1 ' pass /fail
    End Select
        
    For Each oPart In oTree(2).SubTree
        ReDim Preserve vParts(UBound(vParts) + 1)
        vParts(UBound(vParts)) = Expression(oPart)
    Next
    StOr = "(" & Join(vParts, "|") & ")"
End Function

Private Function StRepeat(oTree As ParseTree) As String
    Dim lMin As Long
    Dim lMax As Long
    Dim sModifier As String
    Dim sExpression As String
    
    lMin = -1
    lMax = -1
    
    Select Case oTree(1).Index
        Case 1 ' pass /fail
    End Select
    
    sExpression = Expression(oTree(2))
    
    Select Case Left$(sExpression, 1)
        Case "[", "("
            StRepeat = sExpression
        Case Else
            If Len(sExpression) <> 1 Then
                StRepeat = "(" & sExpression & ")"
            Else
                StRepeat = sExpression
            End If
    End Select

    
    Select Case oTree(4).Index ' min
        Case 1
            lMin = oTree(4).Text
    End Select
    Select Case oTree(5).Index ' max
        Case 1
            lMax = oTree(5).Text
    End Select
    
    If (lMin = -1 Or lMin = 1) And lMax = -1 Then
        sModifier = "+"
    ElseIf lMin = lMax Then
        sModifier = "{" & lMin & "}"
    ElseIf lMin = 0 And lMax = -1 Then
        sModifier = "*"
    ElseIf lMin = 0 And lMax = 1 Then
        sModifier = "?"
    ElseIf lMax = -1 Then
        sModifier = "{" & lMin & ",}"
    ElseIf lMax <> lMin Then
        sModifier = "{" & lMin & "," & lMax & "}"
    End If
    
    StRepeat = StRepeat & sModifier
End Function

Private Function StList(oTree As ParseTree) As String

End Function

Private Function StIn(oTree As ParseTree) As String
    Dim oPart As ParseTree
    Dim bRange(-1 To 256) As Boolean
    Dim lIndex As Long
    Dim bInclude As Boolean
    Dim bCase As Boolean
    Dim lStart As Long
    Dim lEnd As Long
    Dim lTemp As Long
    Dim lCount As Long
    Dim bCheck As Boolean
    Dim bPrevious As Boolean
    
    For Each oPart In oTree(1).SubTree
        Select Case oPart(1).Index
            Case 1 ' not
                bInclude = False
            Case Else
                bInclude = True
        End Select
        
        Select Case oPart(2).Index
            Case 1 ' case
                bCase = True
            Case Else
                bCase = False
        End Select
                
        Select Case oPart(3).Index
            Case 1 ' range
                Select Case oPart(3)(1)(1).Index
                    Case 1 ' number
                        lStart = Val(oPart(3)(1)(1).Text)
                    Case 2 ' string
                        lStart = Asc(oPart(3)(1)(1).Text)
                End Select
                Select Case oPart(3)(1)(2).Index
                    Case 1 ' number
                        lEnd = Val(oPart(3)(1)(2).Text)
                    Case 2 ' string
                        lEnd = Asc(oPart(3)(1)(2).Text)
                End Select
                
                If lEnd < lStart Then
                    lTemp = lEnd
                    lEnd = lStart
                    lStart = lTemp
                End If
                
                For lIndex = lStart To lEnd
                    bRange(lIndex) = bInclude
                Next
            Case 2 ' number
                bRange(Val(oPart(3).Text)) = bInclude
            Case 3 ' string
                For lIndex = 1 To Len(oPart(3).Text)
                    bRange(Asc(Mid$(oPart(3).Text, lIndex, 1))) = bInclude
                Next
        End Select
    Next

    For lIndex = 32 To 255
        If bRange(lIndex) Then
            lCount = lCount + 1
        End If
    Next
    
    If lCount <= 112 Then
        bCheck = True
    Else
        bCheck = False
    End If
    
    bPrevious = bCheck
    lStart = 0
    For lIndex = 0 To 256
        If bRange(lIndex) <> bRange(lIndex - 1) Then
            If bRange(lIndex) = bCheck Then
                lStart = lIndex
            Else
                If (lIndex - 1) > lStart Then
                    StIn = StIn & Chr$(lStart) & "-" & Chr$(lIndex - 1)
                Else
                    StIn = StIn & Chr$(lStart)
                End If
            End If
        End If
    Next
    If bCheck = False Then
        StIn = "[^" & StIn & "]"
    Else
        StIn = "[" & StIn & "]"
    End If
End Function

Private Function StOptional(oTree As ParseTree) As String
    Dim sExpression As String
    
    sExpression = Expression(oTree(2))
    
    Select Case Left$(sExpression, 1)
        Case "[", "("
            StOptional = sExpression & "?"
        Case Else
            If Len(sExpression) <> 1 Then
                StOptional = "(" & sExpression & ")?"
            Else
                StOptional = "(" & sExpression & ")?"
            End If
    End Select
End Function

Private Function StNot(oTree As ParseTree) As String

End Function

Private Function StEach(oTree As ParseTree) As String

End Function

Private Function StEOS(oTree As ParseTree) As String
    StEOS = "$"
End Function

Private Function StBOS(oTree As ParseTree) As String
    StBOS = "^"
End Function

Private Function StPass(oTree As ParseTree) As String

End Function

Private Function StFail(oTree As ParseTree) As String

End Function

Private Function StExternal(oTree As ParseTree) As String

End Function

Private Function StVariable(oTree As ParseTree) As String

End Function

