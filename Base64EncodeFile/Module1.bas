Attribute VB_Name = "Module1"
Option Explicit


Private Const CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Private Type Iterator
    ByteIndex As Long
    ByteBitIndex As Long
    BitIndex As Long
    Value As Long
End Type

Public Function DecodeAndSaveFile(ByVal sPath As String, ByVal sBase64Encoded As String)
    Dim yTarget() As Byte
    
    Dim sFile As String
    sFile = "text.txt"
    
    Base64DecodeString sBase64Encoded, yTarget
    
    If Dir(sPath) <> "" Then
        Kill sPath
    End If
    Open sPath For Binary As #1
    Put #1, , yTarget
    Close #1
End Function

Public Function OpenFileAndEncode(ByVal sFileNamePath As String) As String
    Dim yBytes() As Byte
    
    ReDim yBytes(FileLen(sFileNamePath) - 1)
    Open sFileNamePath For Binary As #1
    Get #1, , yBytes
    Close #1
    
    OpenFileAndEncode = Base64EncodeArray(yBytes)
End Function

Public Function Base64DecodeString(sString As String, yTarget() As Byte)
    Dim yArray() As Byte
    Dim lOffset As Long
    
    If Right$(sString, 2) = "==" Then
        lOffset = 2
        sString = Left$(sString, Len(sString) - 2)
    ElseIf Right$(sString, 1) = "=" Then
        lOffset = 1
        sString = Left$(sString, Len(sString) - 1)
    End If
    
    AsciiToArray sString, yArray
    ReSizeRecord yArray, 8, 6, yTarget
End Function

Public Function Base64EncodeArray(yArray() As Byte) As String
    Dim yTarget() As Byte
    
    ReSizeRecord yArray, 6, 8, yTarget
    Base64EncodeArray = BytesToAscii(yTarget) & String$((((UBound(yArray) + 1) * 8) Mod 6) \ 2, "=")
End Function


Public Function BytesToAscii(yArray() As Byte) As String
    Dim lIndex As Long
    
    BytesToAscii = Space(UBound(yArray) + 1)
    
    For lIndex = 0 To UBound(yArray)
        Mid$(BytesToAscii, lIndex + 1, 1) = Mid$(CHARS, yArray(lIndex) + 1, 1)
    Next
End Function

Public Function AsciiToArray(sString As String, yArray() As Byte)
    Dim lIndex As Long
    
    ReDim yArray(Len(sString) - 1)
    
    For lIndex = 1 To Len(sString)
        yArray(lIndex - 1) = InStr(CHARS, Mid$(sString, lIndex, 1)) - 1
    Next
End Function

Public Sub ReSizeRecord(yData() As Byte, lFromSize As Long, lToSize As Long, yTarget() As Byte)
    Dim itSource As Iterator
    Dim itTarget As Iterator
    Dim lByteIndex As Long
    Dim lFromNoBytes As Long
    Dim lToNoBytes As Long
    Dim lTotalBits As Long
    Dim lMultiplier As Long
    Dim lBitShiftIndex As Long
    Dim lFromMask As Long
    Dim lToMask As Long
    
    lFromNoBytes = (lFromSize - 1) \ 8 + 1
    lToNoBytes = (lToSize - 1) \ 8 + 1
    
    lFromMask = 2 ^ lFromSize - 1
    lToMask = 2 ^ lToSize - 1
    
    lTotalBits = (UBound(yData) + 1) * 8

    For itSource.BitIndex = 0 To lTotalBits - 1 Step lFromSize
        itSource.ByteIndex = itSource.BitIndex \ 8
        itSource.ByteBitIndex = itSource.BitIndex Mod 8
        
        itTarget.ByteIndex = itTarget.BitIndex \ 8
        itTarget.ByteBitIndex = itTarget.BitIndex Mod 8
        
        'read value
        lMultiplier = 1
        itSource.Value = 0
        For lByteIndex = itSource.ByteIndex To itSource.ByteIndex + lFromNoBytes
            If lByteIndex <= UBound(yData) Then
                itSource.Value = itSource.Value + CLng(yData(lByteIndex)) * lMultiplier
            End If
            lMultiplier = lMultiplier * 256
        Next
        For lBitShiftIndex = 0 To itSource.ByteBitIndex - 1
            itSource.Value = itSource.Value \ 2
        Next
    
        itSource.Value = itSource.Value And lFromMask
        
        itTarget.Value = itSource.Value And lToMask
        For lBitShiftIndex = 0 To itTarget.ByteBitIndex - 1
            itTarget.Value = itTarget.Value * 2
        Next
        
        For lByteIndex = itTarget.ByteIndex To itTarget.ByteIndex + lToNoBytes
            ReDim Preserve yTarget(lByteIndex)
            yTarget(lByteIndex) = yTarget(lByteIndex) Or (itTarget.Value And 255)
            itTarget.Value = itTarget.Value \ 256
        Next
        
        itTarget.BitIndex = itTarget.BitIndex + lToSize
    Next
    
    Dim lFromRecords As Long
    Dim lToRecords As Long
    
    
    lFromRecords = lTotalBits \ lFromSize
    lToRecords = Roundup((lFromRecords * lToSize) / 8) - 1
    ReDim Preserve yTarget(lToRecords)
End Sub

Public Function Roundup(dValue As Double) As Long
    If (dValue - Int(dValue)) > 0 Then
        Roundup = Int(dValue) + 1
    Else
        Roundup = Int(dValue)
    End If
End Function

