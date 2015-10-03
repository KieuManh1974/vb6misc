Attribute VB_Name = "Module1"
Option Explicit


Private FragmentSize As Long
Private Const ChunkSize As Long = 65500

Private Type Chunk
    Bytes(1 To ChunkSize) As Byte
End Type

Private oChunk As Chunk

Private Type Info
    Name As String
    Max As Long
End Type

Private aFilesToJoin() As Info

Sub Main()
    JoinFile
End Sub

Private Sub JoinFile()
    Dim oFile As File
    Dim oFSO As New FileSystemObject
    Dim sDefinition As String
    Dim oParser As IParseObject
    Dim oParseTree As ParseTree
    Dim iCharIndex As Integer
    Dim tInfo As Info
    
    sDefinition = sDefinition & "frag_index := AND [CASE '.frg('], {REPEAT IN '0' TO '9'}, [')'], EOS;"
    sDefinition = sDefinition & "scan := AND (REPEAT IN 32 TO 255 UNTIL (OR frag_index, EOS)), frag_index;"
                
    If Not SetNewDefinition(sDefinition) Then
        MsgBox "Bad definition"
        Stop
    End If
    Set oParser = ParserObjects("scan")
    
    ReDim Preserve aFilesToJoin(0) As Info
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        Stream.Text = oFile.Path
        Set oParseTree = New ParseTree
        If oParser.Parse(oParseTree) Then
            tInfo.Name = oParseTree(1).Text
            tInfo.Max = Val(oParseTree(2).Text)
            UpdateFilesToJoin tInfo
        End If
    Next
    
    JoinFiles
End Sub

Private Sub UpdateFilesToJoin(tInfo As Info)
    Dim vFile As Variant
    Dim iFileIndex As Integer
    
    For iFileIndex = 1 To UBound(aFilesToJoin)
        If aFilesToJoin(iFileIndex).Name = tInfo.Name Then
            If tInfo.Max > aFilesToJoin(iFileIndex).Max Then
                aFilesToJoin(iFileIndex).Max = tInfo.Max
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next
    ReDim Preserve aFilesToJoin(UBound(aFilesToJoin) + 1) As Info
    aFilesToJoin(UBound(aFilesToJoin)) = tInfo
End Sub

Private Sub JoinFiles()
    Dim iFileIndex As Integer
    Dim iFragmentIndex As Integer
    Dim sJoinFile As String
    Dim sFragmentFile As String
    Dim lFragmentChunks As Long
    Dim lFragmentRemainder As Long
    Dim lFileSize As Long
    Dim iChunkIndex As Integer
    
    If UBound(aFilesToJoin) > 0 Then
        For iFileIndex = 1 To UBound(aFilesToJoin)
            sJoinFile = aFilesToJoin(iFileIndex).Name
            
            Open sJoinFile For Binary As #1
            
            For iFragmentIndex = 1 To aFilesToJoin(iFileIndex).Max
                sFragmentFile = aFilesToJoin(iFileIndex).Name & ".frg(" & iFragmentIndex & ")"
                lFileSize = FileLen(sFragmentFile)
                lFragmentChunks = lFileSize \ ChunkSize
                lFragmentRemainder = lFileSize Mod ChunkSize
                
                Open sFragmentFile For Binary As #2
                For iChunkIndex = 1 To lFragmentChunks
                    CopyChunk
                Next
                CopyRemainder lFragmentRemainder
                Close #2
                Kill sFragmentFile
            Next
            
            Close #1
        Next
    End If
End Sub

Private Function CopyChunk()
    Get #2, , oChunk
    Put #1, , oChunk
End Function

Private Function CopyRemainder(ByVal lCount As Long)
    Dim iByteIndex As Long
    Dim bByte As Byte
    
    For iByteIndex = 1 To lCount
        Get #2, , bByte
        Put #1, , bByte
    Next
End Function





