Attribute VB_Name = "Combinations"
Option Explicit

Private Type Block
    Position As Long
    Size As Long
End Type

Public Sub TextCombinations()
    Anagram "ABC"
End Sub


Public Sub Anagram(sText)
    Dim lNumber() As Long
    Dim lLen As Long
    Dim sAnagramText As String
    Dim bFinished As Boolean
    Dim lDigitIndex As Long
    Dim lPosition As Long
    
    lLen = Len(sText)
    
    ReDim lNumber(lLen - 1) As Long
    
    While lPosition < lLen
        sAnagramText = sText
        For lDigitIndex = lLen - 2 To 0 Step -1
            sAnagramText = Left$(sAnagramText, lDigitIndex) & Rotate(Mid$(sAnagramText, lDigitIndex + 1), lNumber(lDigitIndex))
            
        Next
        Debug.Print sAnagramText
        
        lPosition = 0
        bFinished = False
        While Not bFinished And lPosition < lLen
            bFinished = True
            lNumber(lPosition) = lNumber(lPosition) + 1
            If lNumber(lPosition) > (lLen - lPosition - 1) Then
                lNumber(lPosition) = 0
                lPosition = lPosition + 1
                bFinished = False
            End If
        Wend
    Wend
End Sub

Private Function Rotate(ByVal sText, ByVal lPosition As Long) As String
    Rotate = Mid$(sText & sText, lPosition + 1, Len(sText))
End Function
