VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnglishNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function EnglishToNumber(ByVal sText As String, Value As String) As Boolean
    Dim oTree As New ParseTree
    
    ParserTextString.ParserText = sText
    If Not oParser.Parse(oTree) Then
        Exit Function
    End If
    
    Value = DecodeNumbers(oTree)
    EnglishToNumber = True
End Function

Private Function DecodeNumbers(oTree As ParseTree) As String
    Dim oNumber As ParseTree
    Dim dMultiplier As Double
    Dim dValue As Double
    
    DecodeNumbers = ""
    For Each oNumber In oTree.SubTree
        dValue = DecodeNumber(oNumber(1))
        DecodeNumbers = DecodeNumbers & dValue
    Next
End Function

Private Function DecodeNumber(oTree As ParseTree) As Double
    Select Case oTree.Index
        Case 1 'million
            DecodeNumber = DecodeMillions(oTree(1))
        Case 2 'thousand
            DecodeNumber = DecodeThousands(oTree(1))
        Case 3 'hundred'
            DecodeNumber = DecodeHundreds(oTree(1))
        Case 4 'ten
            DecodeNumber = DecodeTens(oTree(1))
        Case 5 'unit
            DecodeNumber = DecodeUnits(oTree(1))
        Case 6 'roman
            DecodeNumber = DecodeRomanFigure(oTree(1))
        Case 7 'numeric
    End Select
End Function

Private Function DecodeRomanFigure(oTree As ParseTree) As Double
    ' Thousand
    DecodeRomanFigure = oTree(1).Index * 1000
    
    ' Hundred
    Select Case oTree(2).Index
        Case 1
            DecodeRomanFigure = DecodeRomanFigure + 400
        Case 2
            DecodeRomanFigure = DecodeRomanFigure + 900
        Case 3
            DecodeRomanFigure = DecodeRomanFigure + oTree(2)(1).Index * 100
        Case 4
            DecodeRomanFigure = DecodeRomanFigure + 500
    End Select
    
    ' Ten
    Select Case oTree(3).Index
        Case 1
            DecodeRomanFigure = DecodeRomanFigure + 40
        Case 2
            DecodeRomanFigure = DecodeRomanFigure + 90
        Case 3
            DecodeRomanFigure = DecodeRomanFigure + oTree(3)(1).Index * 10
        Case 4
            DecodeRomanFigure = DecodeRomanFigure + 50
    End Select
    
    ' Unit
    Select Case oTree(4).Index
        Case 1
            DecodeRomanFigure = DecodeRomanFigure + 4
        Case 2
            DecodeRomanFigure = DecodeRomanFigure + 9
        Case 3
            DecodeRomanFigure = DecodeRomanFigure + oTree(4)(1).Index
        Case 4
            DecodeRomanFigure = DecodeRomanFigure + 5
    End Select
End Function

Private Function DecodeMillions(oTree As ParseTree)
    Select Case oTree(1).Index
        Case 4 'unit
            DecodeMillions = 1000000# * DecodeUnit(oTree(1)(1))
        Case 3 'ten
            DecodeMillions = 1000000# * DecodeTens(oTree(1)(1))
        Case 2 'hundred
            DecodeMillions = 1000000# * DecodeHundreds(oTree(1)(1))
        Case 1 'thousand
            DecodeMillions = 1000000# * DecodeHundreds(oTree(1)(1))
    End Select
    
    If oTree(4).Index = 1 Then
        Select Case oTree(4)(1).Index
            Case 4 'unit
                DecodeMillions = DecodeMillions + DecodeUnit(oTree(4)(1)(1))
            Case 3 'ten
                DecodeMillions = DecodeMillions + DecodeTens(oTree(4)(1)(1))
            Case 2 'hundred
                DecodeMillions = DecodeMillions + DecodeHundreds(oTree(4)(1)(1))
            Case 1 'thousand
                DecodeMillions = DecodeMillions + DecodeThousands(oTree(4)(1)(1))
        End Select
    End If
End Function


Private Function DecodeThousands(oTree As ParseTree)
    Select Case oTree(1).Index
        Case 3 'unit
            DecodeThousands = 1000 * DecodeUnit(oTree(1)(1))
        Case 2 'ten
            DecodeThousands = 1000 * DecodeTens(oTree(1)(1))
        Case 1 'hundred
            DecodeThousands = 1000 * DecodeHundreds(oTree(1)(1))
    End Select
    
    If oTree(4).Index = 1 Then
        Select Case oTree(4)(1).Index
            Case 3 'unit
                DecodeThousands = DecodeThousands + DecodeUnit(oTree(4)(1)(1))
            Case 2 'ten
                DecodeThousands = DecodeThousands + DecodeTens(oTree(4)(1)(1))
            Case 1 'hundred
                DecodeThousands = DecodeThousands + DecodeHundreds(oTree(4)(1)(1))
        End Select
    End If
End Function

Private Function DecodeHundreds(oTree As ParseTree) As Double
    Select Case oTree(1).Index
        Case 2 'unit
            DecodeHundreds = 100 * DecodeUnit(oTree(1)(1)(1))
        Case 1 'ten
            DecodeHundreds = 100 * DecodeTens(oTree(1)(1))
    End Select
    
    If oTree(4).Index = 1 Then
        Select Case oTree(4)(1).Index
            Case 2 'unit
                DecodeHundreds = DecodeHundreds + DecodeUnit(oTree(4)(1)(1))
            Case 1 'ten
                DecodeHundreds = DecodeHundreds + DecodeTens(oTree(4)(1)(1))
        End Select
    End If
End Function

Private Function DecodeTens(oTree As ParseTree) As Double
    Select Case oTree.Index
        Case 1 ' teens
            DecodeTens = 10 + oTree(1).Index
        Case 2 ' ten + unit
            DecodeTens = 10 * oTree(1)(1).Index
            If oTree(1)(2).Index = 1 Then
                DecodeTens = DecodeTens + oTree(1)(2)(1).Index
            End If
    End Select
End Function

Private Function DecodeUnits(oTree As ParseTree) As Double
    Select Case oTree.Index
        Case 1
            DecodeUnits = 0
        Case 2
            DecodeUnits = DecodeUnit(oTree(1))
    End Select
End Function

Private Function DecodeUnit(oTree As ParseTree) As Double
    DecodeUnit = oTree.Index
End Function

