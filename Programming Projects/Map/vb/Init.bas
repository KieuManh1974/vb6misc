Attribute VB_Name = "Init"
Option Explicit

Public Evaluator As IParseObject

Sub Main()
    InitialiseParser
    MapPlot.Show
End Sub

Public Sub InitialiseParser()
    Dim definition As String
    Dim oParseTree As New ParseTree
    
    definition = "string := REPEAT IN 'a' TO 'z', 'A' TO 'Z', '|', 122 TO 255;" & _
                "operator := OR '+', '&', '>';" & _
                "level2 := AND '(', expression, ')';" & _
                "follows := AND string,'>',string; " & _
                "level1 := OR follows, string, level2;" & _
                "level0 := AND OPTIONAL '-', level1;" & _
                "expression := LIST level0, operator;" & _
                "expressions := LIST expression, [';'];"
                
    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If
    Set Evaluator = ParserObjects("expressions")
End Sub

Public Function EvalExpression(place_name As String, ByVal oParseTree As ParseTree) As Boolean
    Dim term_index As Long
    Dim operation As Long
    Dim negate As Boolean
    
    operation = 1
    negate = True
    For term_index = 1 To oParseTree.Index * 2 - 1 Step 2
        If term_index > 1 Then
            operation = oParseTree(term_index - 1).Index
        End If
        Select Case operation
            Case 1 ' or
                negate = oParseTree(term_index)(1).Text = "-"
                EvalExpression = EvalExpression Or (EvalLevel1(place_name, oParseTree(term_index)(2)) Xor negate)
                
            Case 2 ' and
                negate = oParseTree(term_index)(1).Text = "-"
                EvalExpression = EvalExpression And (EvalLevel1(place_name, oParseTree(term_index)(2)) Xor negate)
        End Select
    Next
End Function

Public Function EvalLevel1(place_name As String, ByVal oResult As ParseTree) As Boolean
    Select Case oResult.Index
        Case 1 ' follows
            EvalLevel1 = FollowedBy(place_name, oResult(1)(1).Text, oResult(1)(3).Text)
        Case 2 ' string
            If InStr(place_name, oResult.Text) > 0 Then
                EvalLevel1 = True
            End If
        Case 3 'level 2
            EvalLevel1 = EvalExpression(place_name, oResult(1)(2))
    End Select
End Function

Public Function FollowedBy(place_name As String, sText1 As String, sText2 As String) As Boolean
    Dim lPosition1 As Long
    Dim lPosition2  As Long
    
    lPosition1 = InStr(place_name, sText1)
    If lPosition1 = 0 Then
        Exit Function
    End If
    lPosition2 = InStr(place_name, sText2)
    If lPosition2 = 0 Then
        Exit Function
    End If
    
    If lPosition1 < lPosition2 Then
        FollowedBy = True
        Exit Function
    End If
    
    While lPosition2 <= lPosition1 And lPosition2 <> 0
        lPosition2 = InStr(lPosition2 + 1, place_name, sText2)
    Wend
    If lPosition2 > lPosition1 Then
        FollowedBy = True
    End If
End Function
