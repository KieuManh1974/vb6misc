Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Type fiFile
    FileName As String
    Name As String
    Load As Long
    Execution As Long
    Length As Long
    Locked As Boolean
    Side As Long
    HeaderCRC As Long
    DataCRC As Long
    Data() As Byte
    BlockData() As Byte
    BlockLengths() As Long
    BlockDataStarts() As Long
    DriveNumber As Long
    NextFile As String
End Type

Private mfiCatalogue() As fiFile
Private mlCatalogueCount As Long


Private myDiscImage(255, 9, 79, 1) As Byte
Private mlBootOption(2) As Long

Public Function BuildDiscImage(Optional ByVal bInterleaved As Boolean) As Byte()
    Dim lDot As Long
    Dim sName As String
    Dim fiFileInfo As fiFile
    Dim sFileName As String
    Dim sFileDir As String
    Dim yImage() As Byte
    Dim lAddress(1) As Long
    Dim lFileNumber(1) As Long
    Dim lSector As Long
    Dim lCatalogueInfoAdr As Long
    Dim lSideIndex As Long
    Dim lCatalogueIndex As Long
    Dim yInterleavedImage() As Byte
    Dim lDataPosition As Long
    Dim lTrack As Long
    Dim lHiBits As Long
    
    ' Debugging.WriteString "Storage.BuildDiscImage"
    
    ReDim yImage(256& * 10& * 80& - 1, 1)
    
    For lSideIndex = 0 To 1
        ' build catalogue
        yImage(&H100& + 5, lSideIndex) = CatalogueCount(lSideIndex) * 8
        yImage(&H100& + 6, lSideIndex) = (800& And &H300&) \ 256 + mlBootOption(lSideIndex) * 16
        yImage(&H100& + 7, lSideIndex) = (800& And &HFF&)
        
        lAddress(lSideIndex) = &H200&
        
        For lCatalogueIndex = 0 To mlCatalogueCount - 1

            fiFileInfo = mfiCatalogue(lCatalogueIndex)

            If fiFileInfo.DriveNumber = lSideIndex Then
                lDot = InStr(fiFileInfo.Name, ".")
                
                If lDot > 0 Then
                    sFileDir = Left$(fiFileInfo.Name, lDot - 1)
                    sFileName = Mid$(fiFileInfo.Name, lDot + 1)
                Else
                    sFileDir = "$"
                    sFileName = Left$(fiFileInfo.Name, 7)
                End If
                
                sName = "        "
        
                Mid$(sName, 1, Len(sFileName)) = sFileName
                Mid$(sName, 8, 1) = sFileDir
                
                CopyMemory yImage(8 + lFileNumber(lSideIndex) * 8, lSideIndex), ByVal sName, 8&
                yImage(8 + lFileNumber(lSideIndex) * 8 + 7, lSideIndex) = yImage(8 + lFileNumber(lSideIndex) * 8 + 7, lSideIndex) Or -fiFileInfo.Locked * &H80&
                
                lCatalogueInfoAdr = &H100& + 8 + lFileNumber(lSideIndex) * 8
                
                CopyMemory yImage(lCatalogueInfoAdr + 0, lSideIndex), fiFileInfo.Load, 2&
                CopyMemory yImage(lCatalogueInfoAdr + 2, lSideIndex), fiFileInfo.Execution, 2&
                CopyMemory yImage(lCatalogueInfoAdr + 4, lSideIndex), fiFileInfo.Length, 2&
                        
                lSector = lAddress(lSideIndex) \ 256
                lHiBits = (lSector And &H300&) \ 256&
                lHiBits = lHiBits + IIf(fiFileInfo.Load > &H2FFFF, &HC&, 4 * ((fiFileInfo.Load \ &H10000) And 3&))
                lHiBits = lHiBits + ((fiFileInfo.Length \ &H10000) And 3&) * 16
                lHiBits = lHiBits + IIf(fiFileInfo.Execution > &H2FFFF, &HC0&, 64 * ((fiFileInfo.Execution \ &H10000) And 3&))

                yImage(lCatalogueInfoAdr + 6, lSideIndex) = lHiBits
                yImage(lCatalogueInfoAdr + 7, lSideIndex) = lSector And &HFF&
                
                If (lAddress(lSideIndex) + fiFileInfo.Length) >= (256& * 10& * 80&) Then
                    MsgBox "Disc too small"
                    Exit Function
                End If
                
                CopyMemory yImage(lAddress(lSideIndex), lSideIndex), fiFileInfo.Data(0), fiFileInfo.Length
                
                lAddress(lSideIndex) = (lAddress(lSideIndex) + fiFileInfo.Length + 255) And &HFFFFFF00
                
                'Debug.Print (lAddress(lSideIndex) \ 2560) & ":" & (lAddress(lSideIndex) \ 256) Mod 10
                
                lFileNumber(lSideIndex) = lFileNumber(lSideIndex) + 1
            End If
        Next
    Next
    
    Erase lAddress
    If bInterleaved Then
        ReDim yInterleavedImage(256& * 10& * 80& * 2& - 1)
        For lTrack = 0 To 79
            For lSideIndex = 0 To 1
                CopyMemory yInterleavedImage(lDataPosition), yImage(lAddress(lSideIndex), lSideIndex), &HA00&
                lDataPosition = lDataPosition + &HA00&
                lAddress(lSideIndex) = lAddress(lSideIndex) + &HA00&
            Next
        Next
        BuildDiscImage = yInterleavedImage
    Else
        BuildDiscImage = yImage
    End If
End Function

Public Sub AddFile(fiAddFile As fiFile)
    ' Debugging.WriteString "Storage.AddFile"
    
    ReDim Preserve mfiCatalogue(mlCatalogueCount)
    mfiCatalogue(mlCatalogueCount) = fiAddFile
    mlCatalogueCount = mlCatalogueCount + 1
End Sub


Private Function CatalogueCount(ByVal lDriveNumber As Long) As Long
    Dim lFileNumber As Long
    
    ' Debugging.WriteString "Storage.CatalogueCount"
    
    For lFileNumber = 0 To mlCatalogueCount - 1
        If mfiCatalogue(lFileNumber).DriveNumber = lDriveNumber Then
            CatalogueCount = CatalogueCount + 1
        End If
    Next
End Function
