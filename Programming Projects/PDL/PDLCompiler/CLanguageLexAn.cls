VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLanguageLexAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Objs As New Collection
Private program As IParseObject

Friend Sub CreateStructure()
    Dim sAnyRange As String
    Dim iChar As Integer
    
    For iChar = 32 To 255
        sAnyRange = sAnyRange & Chr(iChar)
    Next
    
    With Objs
        ' Single chars
        Create("LeftAngleBracket", "Literal").Initialise "", pcLeave, False, "<"
        Create("RightAngleBracket", "Literal").Initialise "", pcLeave, False, ">"
        Create("LeftRoundBracket", "Literal").Initialise "", pcLeave, False, "("
        Create("RightRoundBracket", "Literal").Initialise "", pcLeave, False, ")"
        Create("LeftSquareBracket", "Literal").Initialise "", pcLeave, False, "["
        Create("RightSquareBracket", "Literal").Initialise "", pcLeave, False, "]"
        Create("LeftBrace", "Literal").Initialise "", pcLeave, False, "{"
        Create("RightBrace", "Literal").Initialise "", pcLeave, False, "}"
        Create("SingleQuote", "Literal").Initialise "", pcLeave, False, "'"
        Create("Comma", "Literal").Initialise "", pcLeave, False, ","
        Create("Semicolon", "Literal").Initialise "", pcLeave, False, ";"
        
        ' Short Keywords
        Create("OptionalShortKeyword", "Literal").Initialise "", pcLeave, False, "?"
        Create("EOSShortKeyword", "Literal").Initialise "", pcLeave, False, "||"
        Create("BOSShortKeyword", "Literal").Initialise "", pcLeave, False, ">>"
        Create("AndShortKeyword", "Literal").Initialise "", pcLeave, False, "&"
        Create("PermShortKeyword", "Literal").Initialise "", pcLeave, False, "$"
        Create("OrShortKeyword", "Literal").Initialise "", pcLeave, False, "|"
        Create("InShortKeyword", "Literal").Initialise "", pcLeave, False, ":"
        Create("NotShortKeyword", "Literal").Initialise "", pcLeave, False, "!"
        Create("RepeatShortKeyword", "Literal").Initialise "", pcLeave, False, "#"
        Create("UntilShortKeyword", "Literal").Initialise "", pcLeave, False, ":"
        Create("MinShortKeyword", "Literal").Initialise "", pcLeave, False, "-"
        Create("MaxShortKeyword", "Literal").Initialise "", pcLeave, False, "+"
        Create("CaseShortKeyword", "Literal").Initialise "", pcLeave, False, "^"
        Create("ToShortKeyword", "Literal").Initialise "", pcLeave, False, "-"
        Create("ListShortKeyword", "Literal").Initialise "", pcLeave, False, "@"
        Create("FailShortKeyword", "Literal").Initialise "", pcLeave, False, "~"
        Create("PassShortKeyword", "Literal").Initialise "", pcLeave, False, "*"
        Create("ExternalShortKeyword", "Literal").Initialise "", pcLeave, False, "="
        Create("EachShortKeyword", "Literal").Initialise "", pcLeave, False, "%"
        Create("SkipShortKeyword", "Literal").Initialise "", pcLeave, False, "£"
        
        ' Long Keywords
        Create("OptionalLongKeyword", "Literal").Initialise "", pcLeave, False, "OPTIONAL", True
        Create("EOSLongKeyword", "Literal").Initialise "", pcLeave, False, "EOS", True
        Create("BOSLongKeyword", "Literal").Initialise "", pcLeave, False, "BOS", True
        Create("AndLongKeyword", "Literal").Initialise "", pcLeave, False, "AND", True
        Create("PermLongKeyword", "Literal").Initialise "", pcLeave, False, "PERM", True
        Create("OrLongKeyword", "Literal").Initialise "", pcLeave, False, "OR", True
        Create("InLongKeyword", "Literal").Initialise "", pcLeave, False, "IN", True
        Create("NotLongKeyword", "Literal").Initialise "", pcLeave, False, "NOT", True
        Create("RepeatLongKeyword", "Literal").Initialise "", pcLeave, False, "REPEAT", True
        Create("UntilLongKeyword", "Literal").Initialise "", pcLeave, False, "UNTIL", True
        Create("MinLongKeyword", "Literal").Initialise "", pcLeave, False, "MIN", True
        Create("MaxLongKeyword", "Literal").Initialise "", pcLeave, False, "MAX", True
        Create("CaseLongKeyword", "Literal").Initialise "", pcLeave, False, "CASE", True
        Create("ToLongKeyword", "Literal").Initialise "", pcLeave, False, "TO", True
        Create("ListLongKeyword", "Literal").Initialise "", pcLeave, False, "LIST", True
        Create("FailLongKeyword", "Literal").Initialise "", pcLeave, False, "FAIL", True
        Create("PassLongKeyword", "Literal").Initialise "", pcLeave, False, "PASS", True
        Create("ExternalLongKeyword", "Literal").Initialise "", pcLeave, False, "EXTERNAL", True
        Create("EachLongKeyword", "Literal").Initialise "", pcLeave, False, "EACH", True
         Create("SkipLongKeyword", "Literal").Initialise "", pcLeave, False, "SKIP", True
         
        ' Variable
        Create("VariableChar", "In").Initialise "", pcLeave, False, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_0123456789"
        Create("VariableStartChar", "In").Initialise "", pcLeave, False, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Create("NotVariableChar", "Not").Initialise "", pcLeave, False, !VariableChar
        Create("RepeatVariableChar", "RepeatUntil").Initialise "", pcLeave, False, !VariableChar, pcLeave, !NotVariableChar, 0, 0
        Create("Variable", "And").Initialise "", pcLeave, False, !VariableStartChar, pcLeave, !RepeatVariableChar, pcLeave
        
        ' Keywords with non-alpha
        Create("OptionalLongKeywordNA", "And").Initialise "", pcLeave, False, !OptionalLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("EOSLongKeywordNA", "And").Initialise "", pcLeave, False, !EOSLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("BOSLongKeywordNA", "And").Initialise "", pcLeave, False, !BOSLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("AndLongKeywordNA", "And").Initialise "", pcLeave, False, !AndLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("PermLongKeywordNA", "And").Initialise "", pcLeave, False, !PermLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("OrLongKeywordNA", "And").Initialise "", pcLeave, False, !OrLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("InLongKeywordNA", "And").Initialise "", pcLeave, False, !InLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("NotLongKeywordNA", "And").Initialise "", pcLeave, False, !NotLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("RepeatLongKeywordNA", "And").Initialise "", pcLeave, False, !RepeatLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("ListLongKeywordNA", "And").Initialise "", pcLeave, False, !ListLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("FailLongKeywordNA", "And").Initialise "", pcLeave, False, !FailLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("PassLongKeywordNA", "And").Initialise "", pcLeave, False, !PassLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("ExternalLongKeywordNA", "And").Initialise "", pcLeave, False, !ExternalLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("EachLongKeywordNA", "And").Initialise "", pcLeave, False, !EachLongKeyword, pcLeave, !NotVariableChar, pcLeave
        Create("SkipLongKeywordNA", "And").Initialise "", pcLeave, False, !SkipLongKeyword, pcLeave, !NotVariableChar, pcLeave
                                
        ' Keywords
        Create("AndKeyword", "Or").Initialise "", pcLeave, False, !AndLongKeywordNA, pcLeave, !AndShortKeyword, pcLeave
        Create("PermKeyword", "Or").Initialise "", pcLeave, False, !PermLongKeywordNA, pcLeave, !PermShortKeyword, pcLeave
        Create("OrKeyword", "Or").Initialise "", pcLeave, False, !OrLongKeywordNA, pcLeave, !OrShortKeyword, pcLeave
        Create("RepeatKeyword", "Or").Initialise "", pcLeave, False, !RepeatLongKeywordNA, pcLeave, !RepeatShortKeyword, pcLeave
        Create("UntilKeyword", "Or").Initialise "", pcLeave, False, !UntilLongKeyword, pcLeave, !UntilShortKeyword, pcLeave
        Create("MinKeyword", "Or").Initialise "", pcLeave, False, !MinLongKeyword, pcLeave, !MinShortKeyword, pcLeave
        Create("MaxKeyword", "Or").Initialise "", pcLeave, False, !MaxLongKeyword, pcLeave, !MaxShortKeyword, pcLeave
        Create("ListKeyword", "Or").Initialise "", pcLeave, False, !ListLongKeywordNA, pcLeave, !ListShortKeyword, pcLeave
        Create("EachKeyword", "Or").Initialise "", pcLeave, False, !EachLongKeywordNA, pcLeave, !EachShortKeyword, pcLeave
        Create("ExternalKeyword", "Or").Initialise "", pcLeave, False, !ExternalLongKeywordNA, pcLeave, !ExternalShortKeyword, pcLeave
        Create("OptionalKeyword", "Or").Initialise "", pcLeave, False, !OptionalLongKeywordNA, pcLeave, !OptionalShortKeyword, pcLeave
        Create("EOSKeyword", "Or").Initialise "", pcLeave, False, !EOSLongKeywordNA, pcLeave, !EOSShortKeyword, pcLeave
        Create("BOSKeyword", "Or").Initialise "", pcLeave, False, !BOSLongKeywordNA, pcLeave, !BOSShortKeyword, pcLeave
        Create("InKeyword", "Or").Initialise "", pcLeave, False, !InLongKeywordNA, pcLeave, !InShortKeyword, pcLeave
        Create("NotKeyword", "Or").Initialise "", pcLeave, False, !NotLongKeywordNA, pcLeave, !NotShortKeyword, pcLeave
        Create("CaseKeyword", "Or").Initialise "", pcLeave, False, !CaseLongKeyword, pcLeave, !CaseShortKeyword, pcLeave
        Create("ToKeyword", "Or").Initialise "", pcLeave, False, !ToLongKeyword, pcLeave, !ToShortKeyword, pcLeave
        Create("FailKeyword", "Or").Initialise "", pcLeave, False, !FailLongKeywordNA, pcLeave, !FailShortKeyword, pcLeave
        Create("PassKeyword", "Or").Initialise "", pcLeave, False, !PassLongKeywordNA, pcLeave, !PassShortKeyword, pcLeave
        Create("SkipKeyword", "Or").Initialise "", pcLeave, False, !SkipLongKeywordNA, pcLeave, !SkipShortKeyword, pcLeave
        Create("Space", "Literal").Initialise "", pcLeave, False, " "
        Create("Blank", "In").Initialise "", pcLeave, False, " " & vbCr & vbLf

        ' Repeat chars
        Create("DoubleSingleQuote", "And").Initialise "", pcLeave, False, !SingleQuote, pcLeave, !SingleQuote, pcLeave
        Create("NotSingleQuote", "Not").Initialise "", pcLeave, False, !SingleQuote
        Create("SingleQuoteNotSingleQuote", "And").Initialise "", pcLeave, False, !SingleQuote, pcLeave, !NotSingleQuote, pcLeave
        Create("Character", "In").Initialise "", pcLeave, False, sAnyRange
        Create("CharacterOrQuote", "Or").Initialise "", pcLeave, False, !DoubleSingleQuote, pcLeave, !Character, pcLeave
        Create("Characters", "RepeatUntil").Initialise "", pcLeave, False, !CharacterOrQuote, pcLeave, !SingleQuoteNotSingleQuote, 0, 0
        Create("NotSpace", "Not").Initialise "", pcLeave, False, !Space
        Create("NotBlank", "Not").Initialise "", pcLeave, False, !Blank
        Create("WhiteSpace", "RepeatUntil").Initialise "", pcOmit, False, !Blank, pcLeave, !NotBlank, 0, 0
                
        ' Number
        Create("NumberChar", "In").Initialise "NumberChar", pcLeave, False, "0123456789"
        Create("NotNumberChar", "Not").Initialise "NotNumberChar", pcLeave, False, !NumberChar
        Create("Figure", "RepeatUntil").Initialise "Figure", pcLeave, False, !NumberChar, pcLeave, !NotNumberChar, 3, 1
        Create("OptionalFigure", "Optional").Initialise "OptionalFigure", pcLeave, False, !Figure, pcLeave
        
        ' IN expressions
        Create("SingleLiteral", "And").Initialise "SingleLiteral", pcLeave, False, !SingleQuote, pcLeave, !Character, pcLeave, !SingleQuote, pcLeave
        Create("FigureOrChar", "Or").Initialise "FigureOrChar", pcLeave, False, !Figure, pcLeave, !SingleLiteral, pcLeave
        Create("RangeStatement", "And").Initialise "RangeStatement", pcLeave, False, !FigureOrChar, pcLeave, !Whitespace, pcLeave, !ToKeyword, pcLeave, !Whitespace, pcLeave, !FigureOrChar, pcLeave
        Create("Literal", "And").Initialise "Literal", pcLeave, False, !SingleQuote, pcLeave, !Characters, pcLeave, !SingleQuote, pcLeave
        Create("InLiteralExp", "Or").Initialise "InLiteralExp", pcLeave, False, !RangeStatement, pcLeave, !Literal, pcLeave, !Figure, pcLeave
        Create("OptionalCase", "Optional").Initialise "OptionalCase", pcLeave, False, !CaseKeyword, pcLeave
        Create("OptionalNot", "Optional").Initialise "OptionalNot", pcLeave, False, !NotKeyword, pcLeave
        Create("InSubExp", "And").Initialise "InSubExp", pcLeave, False, !OptionalNot, pcLeave, !Whitespace, pcLeave, !OptionalCase, pcLeave, !Whitespace, pcLeave, !InLiteralExp, pcLeave
        Create("InCommaSubExp", "And").Initialise "InCommaSubExp", pcLeave, False, !Comma, pcLeave, !Whitespace, pcLeave, !InSubExp, pcLeave
        Create("InRepeatSubExp", "Repeat").Initialise "InRepeatSubExp", pcLeave, False, !InCommaSubExp, pcLeave, 0, 0
        Create("InExpression", "And").Initialise "InExpression", pcLeave, False, !InSubExp, pcLeave, !Whitespace, pcLeave, !InRepeatSubExp, pcLeave

        ' REPEAT UNTIL parts
        Create("MinExpression", "And").Initialise "MinExpression", pcLeave, False, !MinKeyword, pcLeave, !Whitespace, pcLeave, !Figure, pcLeave
        Create("MaxExpression", "And").Initialise "MaxExpression", pcLeave, False, !MaxKeyword, pcLeave, !Whitespace, pcLeave, !Figure, pcLeave
        Create("NotMinExpression", "Not").Initialise "NotMinExpression", pcLeave, False, !MinExpression
        Create("NotMaxExpression", "Not").Initialise "NotMaxExpression", pcLeave, False, !MaxExpression
        Create("OptionalMinExpression", "RepeatUntil").Initialise "OptionalMinExpression", pcLeave, False, !MinExpression, pcLeave, !NotMinExpression, 1, 0
        Create("OptionalMaxExpression", "RepeatUntil").Initialise "OptionalMaxExpression", pcLeave, False, !MaxExpression, pcLeave, !NotMaxExpression, 1, 0
        
        ' Expression parts
        Create "Expression", "Or"
        Create("NotComma", "Not").Initialise "NotComma", pcLeave, False, !Comma
        Create("Plus", "Literal").Initialise "Plus", pcLeave, False, "+"
        Create("NotPlus", "Not").Initialise "NotPlus", pcLeave, False, !Plus
        Create("LiteralExpressionString", "And").Initialise "LiteralExpressionString", pcLeave, False, !SingleQuote, pcLeave, !Characters, pcLeave, !SingleQuote, pcLeave
        Create("LiteralExpression", "Or").Initialise "LiteralExpression", pcLeave, False, !LiteralExpressionString, pcLeave, !Figure, pcLeave
        Create("LiteralListElement", "And").Initialise "LiteralListElement", pcLeave, False, !Whitespace, pcLeave, !Plus, pcLeave, !Whitespace, pcLeave, !LiteralExpression, pcLeave
        Create("NotLiteralListElement", "Not").Initialise "NotLiteralListElement", pcLeave, False, !LiteralListElement
        Create("LiteralListElements", "RepeatUntil").Initialise "LiteralListElements", pcLeave, False, !LiteralListElement, pcLeave, !NotLiteralListElement, 0, 0
        Create("LiteralList", "And").Initialise "LiteralList", pcLeave, False, !OptionalCase, pcLeave, !Whitespace, pcLeave, !LiteralExpression, pcLeave, !LiteralListElements, pcLeave
        
        Create("CommaListElement", "And").Initialise "CommaListElement", pcLeave, False, !Whitespace, pcLeave, !Comma, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave
        Create("CommaListElements", "RepeatUntil").Initialise "CommaListElements", pcLeave, False, !CommaListElement, pcLeave, !NotComma
        Create("ExpressionList", "And").Initialise "ExpressionList", pcLeave, False, !Expression, pcLeave, !CommaListElements, pcLeave
        
        Create("CommaWhitespace", "And").Initialise "CommaWhitespace", pcOmit, False, !Whitespace, pcLeave, !Comma, pcLeave, !Whitespace, pcLeave
        Create("ExpressionList2", "List").Initialise "ExpressionList2", pcLeave, False, !Expression, pcLeave, !CommaWhitespace, pcLeave
        
        ' External function signature
        Create("StringChar", "In").Initialise "StringChar", pcLeave, False, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 "
        Create("String", "Repeat").Initialise "String", pcLeave, False, !StringChar, pcLeave
        Create("ParameterList", "List").Initialise "ParameterList", pcLeave, False, !String, pcLeave, !Comma, pcLeave
        Create("BracketedParameterList", "And").Initialise "BracketedParameterList", pcLeave, False, !LeftRoundBracket, pcLeave, !ParameterList, pcLeave, !RightRoundBracket, pcLeave
        Create("OptionalParameterList", "Optional").Initialise "OptionalParameterList", pcLeave, False, !BracketedParameterList, pcLeave
        Create("FunctionSignature", "And").Initialise "FunctionSignature", pcLeave, False, !Variable, pcLeave, !OptionalParameterList, pcLeave
        
        ' Null control
        Create("PassOrFailKeyword", "Or").Initialise "PassOrFailKeyword", pcLeave, False, !FailKeyword, pcLeave, !PassKeyword, pcLeave
        Create("OptionalPassOrFailKeyword", "Optional").Initialise "OptionalPassOrFailKeyword", pcLeave, False, !PassOrFailKeyword, pcLeave
        
        ' Functions
        Create("AndFunction", "And").Initialise "AndFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !AndKeyword, pcLeave, !Whitespace, pcLeave, !ExpressionList2, pcLeave
        Create("PermFunction", "And").Initialise "PermFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !PermKeyword, pcLeave, !Whitespace, pcLeave, !ExpressionList2, pcLeave
        Create("OrFunction", "And").Initialise "OrFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !OrKeyword, pcLeave, !Whitespace, pcLeave, !ExpressionList2, pcLeave
        Create("RepeatFunction", "And").Initialise "RepeatFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !RepeatKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave, !Whitespace, pcLeave, !OptionalMinExpression, pcLeave, !Whitespace, pcLeave, !OptionalMaxExpression, pcLeave
        Create("RepeatUntilFunction", "And").Initialise "RepeatUntilFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !RepeatKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave, !Whitespace, pcLeave, !UntilKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave, !Whitespace, pcLeave, !OptionalMinExpression, pcLeave, !Whitespace, pcLeave, !OptionalMaxExpression, pcLeave
        Create("ListFunction", "And").Initialise "ListFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !ListKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave, !Whitespace, pcLeave, !Comma, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave, !Whitespace, pcLeave, !OptionalMinExpression, pcLeave, !Whitespace, pcLeave, !OptionalMaxExpression, pcLeave
        Create("EachFunction", "And").Initialise "EachFunction", pcLeave, False, !OptionalPassOrFailKeyword, pcLeave, !Whitespace, pcLeave, !EachKeyword, pcLeave, !Whitespace, pcLeave, !ExpressionList2, pcLeave
        Create("OptionalFunction", "And").Initialise "OptionalFunction", pcLeave, False, !OptionalKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave
        Create("ExternalFunction", "And").Initialise "ExternalFunction", pcLeave, False, !ExternalKeyword, pcLeave, !Whitespace, pcLeave, !FunctionSignature, pcLeave
        Create("InFunction", "And").Initialise "InFunction", pcLeave, False, !InKeyword, pcLeave, !Whitespace, pcLeave, !InExpression, pcLeave
        Create("NotFunction", "And").Initialise "NotFunction", pcLeave, False, !NotKeyword, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave
        Create("SkipFunction", "And").Initialise "SkipFunction", pcLeave, False, !SkipKeyword, pcLeave, !Whitespace, pcLeave, !OptionalFigure, pcLeave
        
        ' Expressions
        Create("ExpressionSub", "Or").Initialise "ExpressionSub", pcLeave, False, !LiteralList, pcLeave, !AndFunction, pcLeave, !PermFunction, pcLeave, !OrFunction, pcLeave, !InFunction, pcLeave, !NotFunction, pcLeave, !RepeatUntilFunction, pcLeave, !RepeatFunction, pcLeave, !OptionalFunction, pcLeave, !ListFunction, pcLeave, !SkipFunction, pcLeave, !EOSKeyword, pcLeave, !BOSKeyword, pcLeave, !PassKeyword, pcLeave, !FailKeyword, pcLeave, !ExternalFunction, pcLeave, !EachFunction, pcLeave, !Variable, pcLeave
        Create("ExpressionIgnore", "And").Initialise "ExpressionIgnore", pcLeave, False, !LeftBrace, pcOmit, !ExpressionSub, pcLeave, !RightBrace, pcOmit
        Create("ExpressionOmit", "And").Initialise "ExpressionOmit", pcLeave, False, !LeftSquareBracket, pcOmit, !ExpressionSub, pcLeave, !RightSquareBracket, pcOmit
        Create("ExpressionOverride", "And").Initialise "ExpressionOverride", pcLeave, False, !LeftAngleBracket, pcOmit, !ExpressionSub, pcLeave, !RightAngleBracket, pcOmit
        Create("ExpressionNormal", "And").Initialise "ExpressionNormal", pcLeave, False, !ExpressionSub, pcLeave
        Create("ExpressionWithBrackets", "And").Initialise "ExpressionWithBrackets", pcLeave, False, !LeftRoundBracket, pcOmit, !ExpressionSub, pcLeave, !RightRoundBracket, pcOmit
        !Expression.Initialise "", pcLeave, False, !ExpressionIgnore, pcLeave, !ExpressionOmit, pcLeave, !ExpressionOverride, pcLeave, !ExpressionWithBrackets, pcLeave, !ExpressionNormal, pcLeave
                        
        ' Statements
        Create("ColonEquals", "Literal").Initialise "ColonEquals", pcOmit, False, ":="
        Create("Definition", "And").Initialise "Definition", pcLeave, False, !Variable, pcLeave, !Whitespace, pcLeave, !ColonEquals, pcLeave, !Whitespace, pcLeave, !Expression, pcLeave
        Create("Statement", "And").Initialise "Statement", pcLeave, False, !Whitespace, pcLeave, !Definition, pcLeave, !Whitespace, pcLeave, !Semicolon, pcLeave, !Whitespace, pcLeave
        Create("Statements", "RepeatUntil").Initialise "Statements", pcLeave, False, !Statement, pcLeave, Create("EndOfFile", "EOS")

        Set program = !Statements
    End With
End Sub

Friend Function Parse(oResult As ParseTree) As Boolean
    Set oResult = New ParseTree
    Parse = program.Parse(oResult)
End Function

Friend Function Create(sName As String, sType) As IParseObject
    Set Create = CreateObject("PDLClasses.P" & sType)
    Objs.Add Create, sName
End Function
