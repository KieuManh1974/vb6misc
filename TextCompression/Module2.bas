Attribute VB_Name = "Module2"
Option Explicit

Public Sub DecompressFile()
    Dim lFileLen As Long
    Dim yBytes() As Byte
    Dim sFile As String
    Dim yReplacementTableSize As Byte
    
    sFile = App.Path & "\compressed.txt"
    lFileLen = FileLen(sFile)
    Open sFile For Binary As #1
    Get #1, , myReplacementTableSize
    If myReplacementTableSize > 0 Then
        ReDim myReplacementTable(2, myReplacementTableSize - 1)
        Get #1, , myReplacementTable
    End If
    ReDim yBytes(lFileLen - 1 - 3 * CLng(myReplacementTableSize) - 1)
    Get #1, , yBytes
    Close #1
    
    Decompress yBytes
    OutputDecompressedFile yBytes
End Sub

Private Sub Decompress(yBytes() As Byte)
    Dim lIndex As Long
    
    For lIndex = myReplacementTableSize - 1 To 0 Step -1
        ReplaceWithPair yBytes, myReplacementTable(1, lIndex), myReplacementTable(2, lIndex), myReplacementTable(0, lIndex)
    Next
End Sub

Private Sub ReplaceWithPair(yBytes() As Byte, yFirstLetter As Byte, ySecondLetter As Byte, yReplacementLetter As Byte)
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim lUbound As Long
    
    lUbound = UBound(yBytes)
    
    While lIndex1 <= lUbound
        If yBytes(lIndex1) = yReplacementLetter Then
            lUbound = lUbound + 1
            ReDim Preserve yBytes(lUbound)
            For lIndex2 = lUbound To lIndex1 + 1 Step -1
                yBytes(lIndex2) = yBytes(lIndex2 - 1)
            Next
            yBytes(lIndex1) = yFirstLetter
            yBytes(lIndex1 + 1) = ySecondLetter
            lIndex1 = lIndex1 + 1
        End If
        lIndex1 = lIndex1 + 1
    Wend
End Sub

Private Sub OutputDecompressedFile(yBytes() As Byte)
    Dim yReplacementTable() As Byte
    Dim sFile As String
    
    sFile = App.Path & "\decompressed.txt"
    If Dir(sFile) <> "" Then
        Kill sFile
    End If
    
    Open sFile For Binary As #1
    Put #1, , yBytes
    Close #1
End Sub
