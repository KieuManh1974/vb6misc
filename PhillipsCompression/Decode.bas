Attribute VB_Name = "Decode"
Option Explicit

Private mdeDictionary() As DictionaryEntry
Private mlDictionarySize As Long
Private Type DictionaryEntry
    Symbol As String * 1
    Chars As String
    Count As Long
End Type
Private msText As String

Public Sub DecodeFile()
    Dim sPath As String
    
    sPath = App.Path & "\compressed.txt"
    
    Open sPath For Binary As #1
    ReadDictionary
    msText = String$(FileLen(sPath) - Seek(1) + 1, " ")
    Get #1, , msText
    Close #1
    
    Decompress
    Debug.Print msText
End Sub

Private Sub ReadDictionary()
    Dim ySymbol As Byte
    Dim ySizeLo As Byte
    Dim ySizeHi As Byte
    Dim sSubstitute As String
    
    Do
'        Get #1, , ySizeLo
'        Get #1, , ySizeHi
'        If ySizeLo = 0 And ySizeHi = 0 Then
'            Exit Do
'        End If
ySizeLo = 2
ySizeHi = 0
        Get #1, , ySymbol
        sSubstitute = String$(ySizeHi * 256 + ySizeLo, " ")
        
        Get #1, , sSubstitute
        
        If ySymbol = 0 And Asc(Left$(sSubstitute, 1)) = 0 And Asc(Right$(sSubstitute, 1)) = 0 Then
            Exit Do
        End If
        
        ReDim Preserve mdeDictionary(mlDictionarySize)
        mdeDictionary(mlDictionarySize).Symbol = Chr$(ySymbol)
        mdeDictionary(mlDictionarySize).Chars = sSubstitute
        'debug.Print Asc(Left$(sSubstitute, 1)) & "/" & Asc(Right$(sSubstitute, 1)) & ":" & ySymbol
        mlDictionarySize = mlDictionarySize + 1
    Loop
End Sub

Private Sub Decompress()
    Dim lIndex As Long
    
    For lIndex = 0 To mlDictionarySize - 1
        msText = Replace$(msText, mdeDictionary(lIndex).Symbol, mdeDictionary(lIndex).Chars)
        'Debug.Print msText
    Next
End Sub
