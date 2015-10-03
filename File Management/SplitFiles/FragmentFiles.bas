Attribute VB_Name = "FragmentFiles"
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Type GUID
    Data(0 To 15) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Public FragmentSize As Long
Private Const ChunkSize As Long = 65500

Private Type Chunk
    Bytes(1 To ChunkSize) As Byte
End Type

Private Type HeaderInfo
    UID As GUID
    Start As Long
    FragmentSize As Long
    FileSize As Long
    Name As String * 50
End Type

Private tChunk As Chunk
Private Const HeaderSize = 16 + 4 + 4 + 4 + 50

Public Sub SplitFilesSub(ByVal FragmentSize As Long)
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
    Dim iByteIndex As Long
    Dim iFragmentIndex As Long

    Dim lFragmentSize As Long
    Dim yByte As Byte
    Dim tHeader As HeaderInfo
    
    lFragmentSize = FragmentSize - HeaderSize
    
    ' No of files
    lFragmentChunks = lFragmentSize \ ChunkSize
    lFragmentRemainder = lFragmentSize Mod ChunkSize
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        If Left$(oFile.Name, 1) <> "#" Then
            sSplitFile = oFile.Path
            lFileSize = oFile.Size
            
            If lFileSize > FragmentSize Then
                lFileFragments = lFileSize \ lFragmentSize
                lFileRemainder = lFileSize Mod lFragmentSize
                
                lFileRemainderChunks = lFileRemainder \ ChunkSize
                lFileRemainderRemainder = lFileRemainder Mod ChunkSize
                
                Open sSplitFile For Binary As #1
                CoCreateGuid tHeader.UID
                With tHeader
                    .FileSize = lFileSize
                    .Name = oFile.Name
                End With
                    
                For iFragmentIndex = 0 To lFileFragments - 1
                    sFragmentFile = sSplitFile & ".frg(" & iFragmentIndex + 1 & ")"
                    If oFSO.FileExists(sFragmentFile) Then
                        Kill sFragmentFile
                    End If
                    Open sFragmentFile For Binary As #2
                    
                    tHeader.Start = iFragmentIndex * lFragmentSize
                    tHeader.FragmentSize = lFragmentSize
                    Put #2, , tHeader
                    
                    For iChunkIndex = 0 To lFragmentChunks - 1
                        Get #1, , tChunk
                        Put #2, , tChunk
                    Next
                    For iByteIndex = 1 To lFragmentRemainder
                        Get #1, , yByte
                        Put #2, , yByte
                    Next
                    Close #2
                Next
                
                If lFileRemainderChunks <> 0 Or lFileRemainderRemainder <> 0 Then
                    sFragmentFile = sSplitFile & ".frg(" & iFragmentIndex + 1 & ")"
                    Open sFragmentFile For Binary As #2
                    tHeader.Start = iFragmentIndex * lFragmentSize
                    tHeader.FragmentSize = lFileRemainder
                    Put #2, , tHeader
                    For iChunkIndex = 0 To lFileRemainderChunks - 1
                        Get #1, , tChunk
                        Put #2, , tChunk
                    Next
                    For iByteIndex = 1 To lFileRemainderRemainder
                        Get #1, , yByte
                        Put #2, , yByte
                    Next
                    Close #2
                End If
    
                Close #1
                Kill sSplitFile
            End If
        End If
    Next
End Sub
