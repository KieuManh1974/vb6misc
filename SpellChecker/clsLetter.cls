VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Link
    Letter As String * 1
    Frequency As Long
    SubLetters As clsLetter
End Type

Private msLetters() As Link
Private mlCount As Long

Public Sub AddWord(ByVal sWord As String)
    Dim lIndex As Long
    Dim sChar As String
    Dim oLetter As clsLetter
    
    Set oLetter = Me
    
    For lIndex = 1 To Len(sWord)
        sChar = Mid$(sWord, lIndex, 1)
        Set oLetter = oLetter.AddLetter(sChar)
    Next
    oLetter.AddLetter Chr$(0)
End Sub

Private Sub RemoveWord(ByVal sWord As String)

End Sub


Public Function WordExists(ByVal sWord As String) As Boolean
    Dim lIndex As Long
    Dim sChar As String
    Dim oLetter As clsLetter
    
    Set oLetter = Me
    sWord = sWord & Chr$(0)
    For lIndex = 1 To Len(sWord)
        sChar = Mid$(sWord, lIndex, 1)
        Set oLetter = oLetter.GetLetter(sChar)
        If oLetter Is Nothing Then
            Exit Function
        End If
    Next
    WordExists = True
End Function

Public Function WordFrequency(ByVal sWord As String) As Long
    Dim lIndex As Long
    Dim sChar As String
    Dim oLetter As clsLetter
    
    Set oLetter = Me
    sWord = sWord & Chr$(0)
    For lIndex = 1 To Len(sWord)
        sChar = Mid$(sWord, lIndex, 1)
        Set oLetter = oLetter.GetLetter(sChar)
        If oLetter Is Nothing Then
            Exit Function
        End If
    Next
End Function


Public Function AddLetter(ByVal sLetter As String) As clsLetter
    Dim lPosition As Long
    Dim lIndex As Long
    
    Dim lStart As Long
    Dim lEnd As Long
    
    lStart = -1
    lEnd = mlCount * 2 - 1
   
    If mlCount > 0 Then
        Do
            lPosition = (lStart + lEnd) \ 2
         
            If sLetter = msLetters(lPosition).Letter Then
                Set AddLetter = msLetters(lPosition).SubLetters
                msLetters(lPosition).Frequency = msLetters(lPosition).Frequency + 1
                Exit Function
            ElseIf sLetter > msLetters(lPosition).Letter Then
                If lPosition + 1 = mlCount Then
                    lPosition = lPosition + 1
                    Exit Do
                End If
                lStart = lPosition
            Else
                If lPosition = 0 Then
                    Exit Do
                End If
                lEnd = lPosition
            End If
            If lEnd - lStart = 1 Then
                lPosition = lEnd
                Exit Do
            End If
        Loop
    End If
    ReDim Preserve msLetters(mlCount)
    
    If mlCount > 0 Then
        For lIndex = mlCount To lPosition + 1 Step -1
            msLetters(lIndex) = msLetters(lIndex - 1)
        Next
    End If
    msLetters(lPosition).Letter = sLetter
    msLetters(lPosition).Frequency = 1
    Set msLetters(lPosition).SubLetters = New clsLetter
    mlCount = mlCount + 1
    Set AddLetter = msLetters(lPosition).SubLetters
End Function

Public Function GetLetter(ByVal sLetter As String) As clsLetter
    Dim lPosition As Long
    Dim lIndex As Long
    
    Dim lStart As Long
    Dim lEnd As Long
    
    lStart = -1
    lEnd = mlCount * 2 - 1
   
    If mlCount > 0 Then
        Do
            lPosition = (lStart + lEnd) \ 2
         
            If sLetter = msLetters(lPosition).Letter Then
                Set GetLetter = msLetters(lPosition).SubLetters
                Exit Function
            ElseIf sLetter > msLetters(lPosition).Letter Then
                If lPosition + 1 = mlCount Then
                    lPosition = lPosition + 1
                    Exit Do
                End If
                lStart = lPosition
            Else
                If lPosition = 0 Then
                    Exit Do
                End If
                lEnd = lPosition
            End If
            If lEnd - lStart = 1 Then
                lPosition = lEnd
                Exit Do
            End If
        Loop
    End If
End Function

Public Function GetLetterFrequency(ByVal sLetter As String) As clsLetter
    Dim lPosition As Long
    Dim lIndex As Long
    
    Dim lStart As Long
    Dim lEnd As Long
    
    lStart = -1
    lEnd = mlCount * 2 - 1
   
    If mlCount > 0 Then
        Do
            lPosition = (lStart + lEnd) \ 2
         
            If sLetter = msLetters(lPosition).Letter Then
                Set GetLetter = msLetters(lPosition).SubLetters
                Exit Function
            ElseIf sLetter > msLetters(lPosition).Letter Then
                If lPosition + 1 = mlCount Then
                    lPosition = lPosition + 1
                    Exit Do
                End If
                lStart = lPosition
            Else
                If lPosition = 0 Then
                    Exit Do
                End If
                lEnd = lPosition
            End If
            If lEnd - lStart = 1 Then
                lPosition = lEnd
                Exit Do
            End If
        Loop
    End If
End Function

Public Function EnumerateWordList(Optional ByVal sWord As String) As String
    Dim lIndex As Long
    
    For lIndex = 0 To mlCount - 1
        If msLetters(lIndex).Letter = Chr$(0) Then
            Debug.Print sWord
            EnumerateWordList = sWord & vbCrLf
        Else
            msLetters(lIndex).SubLetters.EnumerateWordList (sWord & msLetters(lIndex).Letter)
        End If
    Next
End Function

Public Function EnumerateVocab() As String
    Dim lIndex As Long
    For lIndex = 0 To mlCount - 1
        EnumerateVocab = EnumerateVocab & msLetters(lIndex).Letter
    Next
    If mlCount > 0 Then
        EnumerateVocab = EnumerateVocab & "|"
    End If
    For lIndex = 0 To mlCount - 1
        EnumerateVocab = EnumerateVocab & msLetters(lIndex).SubLetters.EnumerateVocab
    Next
End Function

Public Function SetVocab(ByVal sVocab As String, Optional lPosition As Long = 1)
    Dim sChar As String * 1
    Dim sPreviousChar As String
    Dim lIndex As Long
    Dim bEmpty As Boolean
    
    Erase msLetters
    mlCount = 0
    
    Do
        sChar = Mid$(sVocab, lPosition, 1)
        If sChar = "|" Then
            lPosition = lPosition + 1
            Exit Do
        End If
        If mlCount > 0 Then
            ReDim Preserve msLetters(UBound(msLetters) + 1)
        Else
            ReDim msLetters(0)
            bEmpty = False
        End If
        mlCount = mlCount + 1
        With msLetters(UBound(msLetters))
            .Letter = sChar
            Set .SubLetters = New clsLetter
        End With
        sPreviousChar = sChar
        lPosition = lPosition + 1
        If lPosition > Len(sVocab) Then
            Exit Do
        End If
    Loop
    
    If lPosition <= Len(sVocab) Then
        For lIndex = 0 To UBound(msLetters)
            If msLetters(lIndex).Letter <> Chr$(0) Then
                msLetters(lIndex).SubLetters.SetVocab sVocab, lPosition
            End If
        Next
    End If
End Function


