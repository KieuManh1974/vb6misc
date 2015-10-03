Attribute VB_Name = "Module1"
Public oParser As IParseObject

Public Sub InitialiseParser()
    Dim definition As String
    Dim oFSO As New FileSystemObject
        
    definition = oFSO.OpenTextFile(App.Path & "/Number.pdl").ReadAll

    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = ParserObjects("number")
End Sub

Public Function DecodeNumber(sNumber As String) As String
    Dim oTree As New ParseTree
    
    ParserTextString.ParserText = sNumber
    
    If Not oParser.Parse(oTree) Then
        Exit Function
    End If
    
    DecodeNumber = DecodeLevel1(oTree)
End Function

Private Function DecodeLevel1(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1 ' Digits
            DecodeLevel1 = oTree(1).Text
        Case 2 ' Word
            DecodeLevel2 oTree(1)
    End Select
End Function

Private Function DecodeLevel2(oTree As ParseTree) As String
    Dim dValue As Double
    Dim iMultiplierLevel As Integer
    Dim oWord As ParseTree
    
    dValue = 1
    iMultiplierLevel = 0
    For Each oWord In oTree.SubTree
        iMultiplierLevel = oWord(1).Index
        dValue = DecodeWord(LCase$(oWord.Text))
    Next
End Function

Private Function DecodeWord(sWord As String) As Double
    Select Case sWord
        Case "zero"
            DecodeWord = 0
        Case "one"
            DecodeWord = 1
        Case "two"
            DecodeWord = 2
        Case "three"
            DecodeWord = 3
        Case "four"
            DecodeWord = 4
        Case "five"
            DecodeWord = 5
        Case "six"
            DecodeWord = 6
        Case "seven"
            DecodeWord = 7
        Case "eight"
            DecodeWord = 8
        Case "nine"
            DecodeWord = 9
        Case "ten"
            DecodeWord = 10
        Case "eleven"
            DecodeWord = 11
        Case "twelve"
            DecodeWord = 12
        Case "thirteen"
            DecodeWord = 13
        Case "fourteen"
            DecodeWord = 14
        Case "fifteen"
            DecodeWord = 15
        Case "sixteen"
            DecodeWord = 16
        Case "seventeen"
            DecodeWord = 17
        Case "eighteen"
            DecodeWord = 18
        Case "nineteen"
            DecodeWord = 19
        Case "twenty"
            DecodeWord = 20
        Case "thirty"
            DecodeWord = 30
        Case "fourty"
            DecodeWord = 40
        Case "fifty"
            DecodeWord = 50
        Case "sixty"
            DecodeWord = 60
        Case "seventy"
            DecodeWord = 70
        Case "eighty"
            DecodeWord = 80
        Case "ninety"
            DecodeWord = 90
        Case "hundred"
            DecodeWord = 100
        Case "thousand"
            DecodeWord = 1000
        Case "million"
            DecodeWord = 100000
        Case "billion"
            DecodeWord = 1000000000
        Case "trillion"
            DecodeWord = 1000000000000#
    End Select
End Function
