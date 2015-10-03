Attribute VB_Name = "Module1"
Option Explicit

Private FragmentSize As Long
Private Const ChunkSize As Long = 65500

Private Type Chunk
    Bytes(1 To ChunkSize) As Byte
End Type

Private oChunk As Chunk

Sub Main()
    If Val(Command) <> 0 Then
        FragmentSize = Val(Command)
    Else
        'FragmentSize = 1702400 '1.62 MB
        'FragmentSize = 1457664 '1.44 MB
        FragmentSize = 104857600 ' 100 MiB
        'FragmentSize = 736960512 '700 MB
    End If
    SplitFile
End Sub

Private Sub SplitFile()
    Dim oFile As File
    Dim oFSO As New FileSystemObject
    Dim sSplitFile As String
    Dim sFragmentFile As String
    Dim lFileSize As Long
    
    Dim lFragmentChunks As Long
    Dim lFragmentRemainder As Long
    Dim lFileFragments As Long
    Dim lFileRemainder As Long
    Dim lFileRemainderChunks As Long
    Dim lFileRemainderRemainder As Long
    
    Dim iChunkIndex As Long
    Dim iFragmentIndex As Long
    
    lFragmentChunks = FragmentSize \ ChunkSize
    lFragmentRemainder = FragmentSize Mod ChunkSize
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        If UCase$(oFile.Name) <> "SPLITFILES.EXE" And UCase$(oFile.Name) <> "JOINFILES.EXE" And Left$(oFile.Name, 1) <> "#" Then
            sSplitFile = oFile.Path
            lFileSize = oFile.Size
            
            lFileFragments = lFileSize \ FragmentSize
            lFileRemainder = lFileSize Mod FragmentSize
            
            lFileRemainderChunks = lFileRemainder \ ChunkSize
            lFileRemainderRemainder = lFileRemainder Mod ChunkSize
            
            Open sSplitFile For Binary As #1
            
            For iFragmentIndex = 0 To lFileFragments - 1
                sFragmentFile = sSplitFile & ".frg(" & iFragmentIndex + 1 & ")"
                Open sFragmentFile For Binary As #2
                For iChunkIndex = 0 To lFragmentChunks - 1
                    CopyChunk
                Next
                CopyRemainder lFragmentRemainder
                Close #2
            Next
            
            If lFileRemainderChunks <> 0 Or lFileRemainderRemainder <> 0 Then
                sFragmentFile = sSplitFile & ".frg(" & iFragmentIndex + 1 & ")"
                Open sFragmentFile For Binary As #2
                For iChunkIndex = 0 To lFileRemainderChunks - 1
                    CopyChunk
                Next
                CopyRemainder lFileRemainderRemainder
                Close #2
            End If

            Close #1
            Kill sSplitFile
        End If
    Next
End Sub

Private Function CopyChunk()
    Get #1, , oChunk
    Put #2, , oChunk
End Function

Private Function CopyRemainder(ByVal lCount As Long)
    Dim iByteIndex As Long
    Dim bByte As Byte
    
    For iByteIndex = 1 To lCount
        Get #1, , bByte
        Put #2, , bByte
    Next
End Function





