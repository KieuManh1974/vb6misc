Attribute VB_Name = "ReconstructFiles"
Option Explicit

Private Const ChunkSize As Long = 65500

Private Type Chunk
    Bytes(1 To ChunkSize) As Byte
End Type

Private tChunk As Chunk

Private Type GUID
    Data(0 To 15) As Byte
End Type

Private Type HeaderInfo
    UID As GUID
    Start As Long
    FragmentSize As Long
    FileSize As Long
    Name As String * 50
End Type

Private Type Info
    Header As HeaderInfo
    Filename As String
    Valid As Boolean
End Type

Private Const HeaderSize = 16 + 4 + 4 + 4 + 50

Private aFragmentList() As Info

Public Sub JoinFilesSub()
    ReadFragments
    SortFragments
    RebuildFragments
    RenameFiles
End Sub

Private Sub ReadFragments()
    Dim oFile As File
    Dim oFSO As New FileSystemObject
    Dim tInfo As Info
    
    ReDim aFragmentList(0) As Info

    For Each oFile In oFSO.GetFolder(App.Path).Files
        Open oFile.Path For Binary As #1
        If oFile.Size > HeaderSize Then
            Get #1, , tInfo.Header
            tInfo.Filename = oFile.Path
            ReDim Preserve aFragmentList(UBound(aFragmentList) + 1) As Info
            aFragmentList(UBound(aFragmentList)) = tInfo
        End If
        Close #1
    Next
End Sub

Private Sub SortFragments()
    Dim bSorted As Boolean
    Dim iIndex As Integer
    Dim tTempInfo As Info
    
    While Not bSorted
        bSorted = True
        For iIndex = 1 To UBound(aFragmentList) - 1
            Select Case CompareGUID(aFragmentList(iIndex).Header.UID, aFragmentList(iIndex + 1).Header.UID)
                Case -1 ' First lower
                Case 0 ' Same
                    If aFragmentList(iIndex).Header.Start > aFragmentList(iIndex + 1).Header.Start Then
                        tTempInfo = aFragmentList(iIndex)
                        aFragmentList(iIndex) = aFragmentList(iIndex + 1)
                        aFragmentList(iIndex + 1) = tTempInfo
                        bSorted = False
                    End If
                Case 1 ' First higher
                    tTempInfo = aFragmentList(iIndex)
                    aFragmentList(iIndex) = aFragmentList(iIndex + 1)
                    aFragmentList(iIndex + 1) = tTempInfo
                    bSorted = False
            End Select
        Next
    Wend
End Sub

Private Function CompareGUID(tGUID1 As GUID, tGUID2 As GUID) As Integer
    Dim iIndex As Integer
    
    For iIndex = 0 To 15
        If tGUID1.Data(iIndex) > tGUID2.Data(iIndex) Then
            CompareGUID = 1
            Exit Function
        ElseIf tGUID1.Data(iIndex) < tGUID2.Data(iIndex) Then
            CompareGUID = -1
            Exit Function
        End If
    Next
End Function

Private Sub RebuildFragments()
    Dim sCurrentFile As String
    Dim iIndex As Integer
    Dim tTemp As Info
    Dim iChunks As Long
    Dim iRemainder As Long
    Dim iChunkIndex As Long
    
    Dim yByte As Byte
    
    RemoveInvalidFiles
    
    iIndex = 1
    While iIndex < UBound(aFragmentList)
        If CompareGUID(aFragmentList(iIndex).Header.UID, aFragmentList(iIndex + 1).Header.UID) = 0 Then
            If aFragmentList(iIndex + 1).Header.Start = (CDec(aFragmentList(iIndex).Header.Start) + CDec(aFragmentList(iIndex).Header.FragmentSize)) And aFragmentList(iIndex).Header.FragmentSize <> 0 Then
                MergeFiles aFragmentList(iIndex), aFragmentList(iIndex + 1)
                aFragmentList(iIndex).Header.FragmentSize = aFragmentList(iIndex).Header.FragmentSize + aFragmentList(iIndex + 1).Header.FragmentSize
                RemoveFile iIndex + 1
            Else
                iIndex = iIndex + 1
            End If
        Else
            iIndex = iIndex + 1
        End If
    Wend
End Sub

Private Sub RenameFiles()
    Dim iChunks As Long
    Dim iRemainder As Long
    Dim yByte As Byte
    Dim iChunkIndex As Long
    Dim iIndex As Long
    
    For iIndex = 1 To UBound(aFragmentList)
        If aFragmentList(iIndex).Header.FragmentSize = aFragmentList(iIndex).Header.FileSize Then
            Open aFragmentList(iIndex).Filename For Binary As #2
            Open App.Path & "\" & aFragmentList(iIndex).Header.Name For Binary As #1
            Seek #2, HeaderSize + 1
            iChunks = (aFragmentList(iIndex).Header.FileSize) \ ChunkSize
            iRemainder = (aFragmentList(iIndex).Header.FileSize) Mod ChunkSize
            For iChunkIndex = 1 To iChunks
                Get #2, , tChunk
                Put #1, , tChunk
            Next
            For iChunkIndex = 1 To iRemainder
                Get #2, , yByte
                Put #1, , yByte
            Next
            Close #2
            Close #1
            Kill aFragmentList(iIndex).Filename
        End If
    Next
End Sub

Private Sub MergeFiles(tInfo1 As Info, tInfo2 As Info)
    Dim iChunkIndex As Long
    Dim iChunks As Long
    Dim iRemainder As Long
    Dim yByte As Byte
    
    Open tInfo1.Filename For Binary As #1
    Open tInfo2.Filename For Binary As #2
    
    Seek #1, 16 + 4 + 1
    Put #1, , CLng(tInfo1.Header.FragmentSize + tInfo2.Header.FragmentSize)
    Seek #1, tInfo1.Header.FragmentSize + HeaderSize + 1
    Seek #2, HeaderSize + 1
    
    iChunks = (tInfo2.Header.FragmentSize) \ ChunkSize
    iRemainder = (tInfo2.Header.FragmentSize) Mod ChunkSize
    For iChunkIndex = 1 To iChunks
        Get #2, , tChunk
        Put #1, , tChunk
    Next
    For iChunkIndex = 1 To iRemainder
        Get #2, , yByte
        Put #1, , yByte
    Next
    Close #2
    Close #1
    
    Kill tInfo2.Filename
End Sub

Private Sub RemoveInvalidFiles()
    Dim iIndex As Long
    Dim iIndex2 As Long
    
    For iIndex = 1 To UBound(aFragmentList) - 1
        If CompareGUID(aFragmentList(iIndex).Header.UID, aFragmentList(iIndex + 1).Header.UID) = 0 Then
            aFragmentList(iIndex).Valid = True
            aFragmentList(iIndex + 1).Valid = True
        End If
    Next
    
    iIndex = 1
    While iIndex <= UBound(aFragmentList)
        If Not aFragmentList(iIndex).Valid Then
            RemoveFile iIndex
        Else
            iIndex = iIndex + 1
        End If
    Wend
End Sub

Private Sub RemoveFile(iIndex As Long)
    Dim iIndex2 As Long
    
    For iIndex2 = iIndex To UBound(aFragmentList) - 1
        aFragmentList(iIndex2) = aFragmentList(iIndex2 + 1)
    Next
    ReDim Preserve aFragmentList(UBound(aFragmentList) - 1) As Info
End Sub

