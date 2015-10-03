Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()
    CountWays 5
End Sub

Public Sub CountWays(ByVal lNumber As Long)
    Dim lCount(100) As Long
    Dim lDens(5) As Long
    Dim lCipher(5) As Long
    Dim bFinished As Boolean
    Dim lValue As Long
    Dim lIndex As Long
    Dim lAddend As Long
    Dim lNewDigit As Long
    
    lDens(0) = -5
    lDens(1) = -2
    lDens(2) = -1
    lDens(3) = 1
    lDens(4) = 2
    lDens(5) = 5
   ' lDens(6) = 10
    
    Do
        lValue = 0
        For lIndex = 0 To 5
            lValue = lValue + lDens(lIndex) * lCipher(lIndex)
        Next
        If lValue <= 100 And lValue >= 0 Then
            lCount(lValue) = lCount(lValue) + 1
        End If
        
        lIndex = 0
        lAddend = 1
        While lAddend <> 0 And lIndex < 6
            lNewDigit = lCipher(lIndex) + lAddend
            lAddend = lNewDigit \ lNumber
            lCipher(lIndex) = lNewDigit Mod lNumber
            lIndex = lIndex + 1
        Wend
        If lIndex = 6 And lAddend = 1 Then
            bFinished = True
        End If
    Loop Until bFinished
End Sub
