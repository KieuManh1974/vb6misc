Attribute VB_Name = "Module1"
Option Explicit

Private lSpellChecker As Long
Private Const SYMBOL_COUNT = 4
Private moVocab As New clsLetter

Public Sub Main()
    Dim sFile As String
    Dim vWords As Variant
    Dim vWord As Variant
    
    sFile = String$(FileLen(App.Path & "\wordlist.txt"), Chr$(0))
    Open App.Path & "\wordlist.txt" For Binary As #1
    Get #1, , sFile
    vWords = Split(sFile, vbCrLf)
    
    moVocab.AddWord "when"
    moVocab.AddWord "where"
    
    For Each vWord In vWords
        moVocab.AddWord UCase$(vWord)
    Next
    
    Debug.Print moVocab.EnumerateWordList
    Debug.Print moVocab.EnumerateVocab
End Sub
