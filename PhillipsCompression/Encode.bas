Attribute VB_Name = "Encode"
Option Explicit

Private Type Pair
    Chars As String
    Frequency As Long
End Type
    
Private Type DictionaryEntry
    Symbol As String * 1
    Chars As String
    Count As Long
End Type

Private sText As String
Private mlCharCount(255) As Long
Private mpaCounts() As Pair
Private mlSubstitueSymbol As Long
Private mbEscapedSubstituteSymbol As Boolean
Private mdeDictionary() As DictionaryEntry
Private mlDictionarySize As Long
Private msMostCommonPair As String * 2
Private mlHighestFrequencyPair As Long

Sub main()
    EncodeFile
End Sub

Private Sub EncodeFile()
    Dim sPath As String
    
    sPath = App.Path & "\uncompressed.txt"
    sText = String$(FileLen(sPath), " ")
    Open sPath For Binary As #1
    Get #1, , sText
    Close #1

    'sText = "The most astonishing thing I ever saw, was in the depths of those catacombs.  I often wondered what it would be like to be blind."
    Debug.Print Len(sText)
        
    Do
        CountSingle
        FindLeastUsedSymbol
        If mbEscapedSubstituteSymbol Then
            Exit Do
        End If
        CountPairs False
        
        'If mlHighestFrequencyPair > 5 Then
        If mlHighestFrequencyPair > 3 Then
            ReDim Preserve mdeDictionary(mlDictionarySize)
            mdeDictionary(mlDictionarySize).Chars = msMostCommonPair
            mdeDictionary(mlDictionarySize).Symbol = Chr$(mlSubstitueSymbol)
            mdeDictionary(mlDictionarySize).Count = mlHighestFrequencyPair
            SubstituteDictionaryEntry
            mlDictionarySize = mlDictionarySize + 1
            sText = Replace$(sText, msMostCommonPair, Chr$(mlSubstitueSymbol))
            Debug.Print Len(sText) & ":" & mlHighestFrequencyPair & ":" & msMostCommonPair & ":" & Asc(Left$(msMostCommonPair, 1)) & "/" & Asc(Right$(msMostCommonPair, 1)) & ":" & mlSubstitueSymbol
        Else
            Exit Do
        End If
    Loop
    WriteFile
    Decode.DecodeFile
End Sub

Private Sub SubstituteDictionaryEntry()
    Dim lIndex As Long
    Dim sSymbols As String
    Dim vSubSymbol As Variant
    Dim sNewSubstitute As String
    Dim sSubSymbol As String * 1
    
    sSymbols = mdeDictionary(mlDictionarySize).Chars
    
    For lIndex = 1 To Len(sSymbols)
        sSubSymbol = Mid$(sSymbols, lIndex, 1)
        vSubSymbol = FindDictionaryEntryFrequency(sSubSymbol)
        If vSubSymbol(1) <> -1 Then
'            If vSubSymbol(1) = mdeDictionary(mlDictionarySize).Count Then
'                Debug.Print vSubSymbol(0)
'                sNewSubstitute = sNewSubstitute & mdeDictionary(vSubSymbol(0)).Chars
'                RemoveDictionaryEntry vSubSymbol(0)
'            Else
'                sNewSubstitute = sNewSubstitute & sSubSymbol
'            End If
            sNewSubstitute = sNewSubstitute & sSubSymbol
        Else
            sNewSubstitute = sNewSubstitute & sSubSymbol
        End If
    Next
    mdeDictionary(mlDictionarySize).Chars = sNewSubstitute
End Sub

Private Sub RemoveDictionaryEntry(ByVal lEntryIndex As Long)
    Dim lIndex As Long
    
    For lIndex = lEntryIndex To mlDictionarySize - 1
        mdeDictionary(lIndex) = mdeDictionary(lIndex + 1)
    Next
    mlDictionarySize = mlDictionarySize - 1
    If mlDictionarySize = -1 Then
        Erase mdeDictionary
    Else
        ReDim Preserve mdeDictionary(mlDictionarySize)
    End If
End Sub

Private Function FindDictionaryEntryFrequency(ByVal sSymbol As String) As Variant
    Dim lFind As Long
    
    For lFind = mlDictionarySize - 1 To 0 Step -1
        If mdeDictionary(lFind).Symbol = sSymbol Then
            FindDictionaryEntryFrequency = Array(lFind, mdeDictionary(lFind).Count)
            Exit Function
        End If
    Next
    FindDictionaryEntryFrequency = Array("", -1)
End Function

Private Sub CountPairs(ByVal bEscaped As Boolean)
    Dim lIndex As Long
    Dim sPair As String * 2
    Dim lCount As Long
    Dim lCheck As Long
    Dim bFound As Boolean
    Dim lCountWithHighestFrequency As Long
    
    Erase mpaCounts
    
    lIndex = 1
    mlHighestFrequencyPair = -1
    
    While lIndex < Len(sText)
        sPair = Mid$(sText, lIndex, 2)
    
        bFound = False
        For lCheck = 0 To lCount - 1
            If mpaCounts(lCheck).Chars = sPair Then
                mpaCounts(lCheck).Frequency = mpaCounts(lCheck).Frequency + 1
                If mpaCounts(lCheck).Frequency > mlHighestFrequencyPair Then
                    mlHighestFrequencyPair = mpaCounts(lCheck).Frequency
                    lCountWithHighestFrequency = lCheck
                End If
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            ReDim Preserve mpaCounts(lCount)
            mpaCounts(lCount).Chars = sPair
            mpaCounts(lCount).Frequency = 1
            If mpaCounts(lCheck).Frequency > mlHighestFrequencyPair Then
                mlHighestFrequencyPair = mpaCounts(lCheck).Frequency
                lCountWithHighestFrequency = lCheck
            End If
            lCount = lCount + 1
        End If
        
        If Mid$(sPair, 1, 1) = Mid$(sPair, 2, 1) Then
            lIndex = lIndex + 2
        Else
            lIndex = lIndex + 1
        End If
    Wend
    msMostCommonPair = mpaCounts(lCountWithHighestFrequency).Chars
End Sub

Private Sub CountSingle()
    Dim lIndex As Long
    Dim lChar As Long
    
    Erase mlCharCount
    
    For lIndex = 1 To Len(sText)
        lChar = Asc(Mid$(sText, lIndex, 1))
        mlCharCount(lChar) = mlCharCount(lChar) + 1
    Next
End Sub

Private Sub FindLeastUsedSymbol()
    Dim lLowestFrequency As Long
    Dim lLowestFrequencySymbol As Long
    Dim lCheck As Long
    Dim lCheckIndex As Long
    
    lLowestFrequency = Len(sText) + 1
    
    For lCheck = 0 To 255
        lCheckIndex = (lCheck + 32) Mod 256
        If mlCharCount(lCheckIndex) < lLowestFrequency Then
            lLowestFrequency = mlCharCount(lCheckIndex)
            lLowestFrequencySymbol = lCheckIndex
            If lLowestFrequency = 0 Then
                Exit For
            End If
        End If
    Next
    
    mlSubstitueSymbol = lLowestFrequencySymbol
    mbEscapedSubstituteSymbol = lLowestFrequency > 0
End Sub

Private Sub WriteFile()
    Dim lIndex As Long
    
    If Dir(App.Path & "\compressed.txt") <> "" Then
        Kill App.Path & "\compressed.txt"
    End If
    Open App.Path & "\compressed.txt" For Binary As #1
    For lIndex = mlDictionarySize - 1 To 0 Step -1
        'Put #1, , CByte(Len(mdeDictionary(lIndex).Chars) And &HFF&)
        'Put #1, , CByte((Len(mdeDictionary(lIndex).Chars) \ 256) And &HFF&)
        Put #1, , CByte(Asc(mdeDictionary(lIndex).Symbol))
        Put #1, , mdeDictionary(lIndex).Chars
    Next
    Put #1, , CByte(0)
    Put #1, , CByte(0)
    Put #1, , CByte(0)
    For lIndex = 1 To Len(sText)
        Put #1, , CByte(Asc(Mid$(sText, lIndex, 1)))
    Next
    Close #1
End Sub
