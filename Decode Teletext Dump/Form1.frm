VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private myFile() As Byte

Dim myPages(39, 31, 15, &H8FF&) As Byte
Dim mbPageLoaded(&H8FF&, 15) As Boolean
Dim mlPageRow(&H8FF&) As Long

Private Sub Form_Load()
    Dim yPage() As Byte
    Dim lPageSize As Long
    
    Dim fiPage As fiFile
    Dim yDisc() As Byte
    Dim sDiscFilename As String
    
    Dim lPage As Long
    Dim lSubPage As Long
    Dim sPage As String
    Dim sSubPage As String
    
    Dim lTotalFiles As Long
    
    Dim yIndex() As Byte
    Dim lIndexSize As Long
    
    Dim yFileData() As Byte
    Dim bHasFileData As Boolean
    
    Dim lDriveNumber As Long
    Dim lPageAddress As Long
    Dim yBlankLineMask() As Byte
    
    Dim lMaxDiscSize As Long
    Dim lMaxFileSize As Long
    Dim lFileLoadAddress As Long
    Dim lTotalFilesSize As Long
    Dim bDiscSizeExceeded As Boolean
    Dim bFileSizeExceeded As Boolean
    
    lMaxDiscSize = 80& * 10& * 256&
    lTotalFilesSize = 512
    
    lFileLoadAddress = &H1F90&
    lMaxFileSize = &H7C00& - lFileLoadAddress
    
    DecodeTeletextFile
    
    For lPage = Page("000") To Page("8FF")
        sPage = HexNum(lPage, 3)
        For lSubPage = 0 To 15
            sSubPage = HexNum(lSubPage, 1)
            If mbPageLoaded(lPage, lSubPage) Then
                yPage = GetPage(lPage, lSubPage)
                
                yBlankLineMask = RemoveBlankLines(yPage)
                CompressData yPage
                AppendMask yPage, yBlankLineMask
                
                lPageSize = UBound(yPage) + 1
                
                bDiscSizeExceeded = False
                bFileSizeExceeded = False
                
                If (AdjustedFileSize(lIndexSize) * 3 + lTotalFilesSize + lPageAddress + lPageSize) >= lMaxDiscSize Then
                    bDiscSizeExceeded = True
                    Debug.Print HexNum(lPage, 3)
                End If
                
                If (lPageAddress + lPageSize) >= lMaxFileSize Then
                    bFileSizeExceeded = True
                End If
                
                If bFileSizeExceeded Or bDiscSizeExceeded Then
                    With fiPage
                        .DriveNumber = lDriveNumber
                        .Name = Format$(lTotalFiles, "00")
                        .Load = lFileLoadAddress
                        .Execution = .Load
                        .Length = UBound(yFileData) + 1
                        .Data = yFileData
                    End With
                    
                    AddFile fiPage
                    lTotalFiles = lTotalFiles + 1
                    Erase yFileData
                    bHasFileData = False
                    
                    If bDiscSizeExceeded Then
                        lTotalFilesSize = 512
                        lDriveNumber = lDriveNumber + 1
                    Else
                        lTotalFilesSize = lTotalFilesSize + AdjustedFileSize(lPageAddress)
                    End If
                    
                    lPageAddress = 0
                Else
                
                    ReDim Preserve yFileData(lPageAddress + lPageSize - 1)
                    CopyMemory yFileData(lPageAddress), yPage(0), lPageSize
                    bHasFileData = True
    
                    ReDim Preserve yIndex(lIndexSize + 7)
                    yIndex(lIndexSize) = Asc(Mid$(sPage, 1, 1))
                    yIndex(lIndexSize + 1) = Asc(Mid$(sPage, 2, 1))
                    yIndex(lIndexSize + 2) = Asc(Mid$(sPage, 3, 1))
                    yIndex(lIndexSize + 3) = Asc(Left$(Format$(lTotalFiles, "00"), 1))
                    yIndex(lIndexSize + 4) = Asc(Right$(Format$(lTotalFiles, "00"), 1))
                    yIndex(lIndexSize + 5) = lPageAddress And &HFF&
                    yIndex(lIndexSize + 6) = lPageAddress \ &H100&
                    yIndex(lIndexSize + 7) = Asc(CStr(lDriveNumber * 2))
                    
                    lIndexSize = lIndexSize + 8
                    
                    lPageAddress = lPageAddress + lPageSize
                End If
            End If
        Next
    Next
    
    
    If bHasFileData Then
        With fiPage
            .DriveNumber = lDriveNumber
            .Name = Format$(lTotalFiles, "00")
            .Load = lFileLoadAddress
            .Execution = .Load
            .Length = UBound(yFileData) + 1
            .Data = yFileData
        End With
        
        AddFile fiPage
    End If
    
    With fiPage
        .DriveNumber = 0
        .Name = "INDEX"
        .Load = &H7C00&
        .Execution = .Load
        .Length = lIndexSize
        .Data = yIndex
    End With
    
    AddFile fiPage
    
    yDisc = BuildDiscImage(True)
    
    sDiscFilename = App.Path & "\teletext.dsd"
    If Dir(sDiscFilename) <> "" Then
        Kill sDiscFilename
    End If
    Open sDiscFilename For Binary As #1
    Put #1, , yDisc
    Close #1
End Sub

Private Function AdjustedFileSize(ByVal lSize As Long)
    AdjustedFileSize = ((lSize + 256&) \ 256) * 256
End Function

Public Sub DecodeTeletextFile()
    Dim lIndex As Long
    Dim vPacket As Variant
    Dim sFile As String
    Dim lMagazine(8) As Long
    Dim lMagazineSubPage(8) As Long
    Dim lColumn As Long
    
    Dim lThisMagazine As Long
    Dim lRow As Long
    Dim lThisMagazinePage As Long
    Dim lSubPage As Long
    
    Dim lChar As Long
    Dim lTempColumn As Long
    
    Dim lValue1 As Long
    Dim lValue2 As Long
    
    sFile = App.Path & "\teletext.txt"
    ReDim myFile(FileLen(sFile) - 1)
    Open sFile For Binary As #1
    Get #1, , myFile
    Close #1
    
    For lIndex = 0 To UBound(myFile)
        If myFile(lIndex) = &H98& Then
            vPacket = DecodePacket(lIndex + 1)
            If vPacket(0) <> -1 Then
                lIndex = lIndex + 40
                
                lThisMagazine = vPacket(0)
                lRow = vPacket(1)
                
                Select Case lRow
                    Case 0
                        lSubPage = vPacket(2)(2)
                        lThisMagazinePage = vPacket(2)(1) * 16 + vPacket(2)(0)
                        lMagazine(lThisMagazine) = lThisMagazine * 256 + lThisMagazinePage
                        
                        If lSubPage > 0 Then
                            lMagazineSubPage(lThisMagazine) = lSubPage - 1
                        Else
                            lMagazineSubPage(lThisMagazine) = 0
                        End If
                        
                        mbPageLoaded(lMagazine(lThisMagazine), lMagazineSubPage(lThisMagazine)) = True
                        
                        
                        If lThisMagazine = 1 Then
                            If lMagazine(1) = ConvertBase("100", 16) Then
                                Debug.Print HexNum(vPacket(2)(2), 1) & HexNum(vPacket(2)(3), 1) & HexNum(vPacket(2)(4), 1) & HexNum(vPacket(2)(5), 1)
                                'Stop
                            End If
                        End If
        
                        mlPageRow(lMagazine(lThisMagazine)) = 0
                        
'                        If lThisMagazine = 1 Then
'                            If lMagazine(1) = ConvertBase("100", 16) Then
'                                Debug.Print
'                                Debug.Print "(" & lRow & ")";
'                            End If
'                        End If
                        For lColumn = 0 To 39
                            myPages(lColumn, lRow, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = vPacket(2)(lColumn)
                            
'                            If lThisMagazine = 1 Then
'                                If lMagazine(1) = ConvertBase("100", 16) Then
'                                    Debug.Print Chr$(vPacket(2)(lColumn));
'                                End If
'                            End If
                        Next
                    Case 1 To 23
'                        If lThisMagazine = 1 Then
'                            If lMagazine(1) = ConvertBase("100", 16) Then
'                                Debug.Print
'                                Debug.Print "(" & lRow & ")";
'                            End If
'                        End If
                        If lRow > mlPageRow(lMagazine(lThisMagazine)) Then
                            For lColumn = 0 To 39
                                myPages(lColumn, lRow, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = vPacket(2)(lColumn)
'                                If lThisMagazine = 1 Then
'                                    If lMagazine(1) = ConvertBase("100", 16) Then
'                                        Debug.Print Chr$(vPacket(2)(lColumn));
'                                    End If
'                                End If
                            Next
                            mlPageRow(lMagazine(lThisMagazine)) = lRow
                        Else
                            mlPageRow(lMagazine(lThisMagazine)) = 99 ' invalidate following rows (corrupted header)
                        End If
                    Case 24
                        If lRow > mlPageRow(lMagazine(lThisMagazine)) Then
                            mlPageRow(lMagazine(lThisMagazine)) = lRow
                        Else
                            mlPageRow(lMagazine(lThisMagazine)) = 99 ' invalidate following rows (corrupted header)
                        End If
                    Case 25 ' replacement
                    
                    Case 27
                        For lChar = 0 To 39
                            lValue1 = InverseHamming(vPacket(2)(lChar))
                            If lValue1 <> -1 Then
                                vPacket(2)(lChar) = InverseHamming(vPacket(2)(lChar))
                            Else
                                vPacket(2)(lChar) = 0
                            End If
                        Next
                        For lChar = 1 To 39 Step 6
                            'Debug.Print lThisMagazine & HexNum(vPacket(2)(lChar + 1), 1) & HexNum(vPacket(2)(lChar), 1)
                            If lChar < 25 Then
                                lTempColumn = 10 * ((lChar - 1) \ 6)
                                myPages(lTempColumn, 24, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = 129 + lTempColumn \ 10
                                myPages(lTempColumn + 1, 24, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = lThisMagazine + 48
                                myPages(lTempColumn + 2, 24, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = Asc(HexNum(vPacket(2)(lChar + 1), 1))
                                myPages(lTempColumn + 3, 24, lMagazineSubPage(lThisMagazine), lMagazine(lThisMagazine)) = Asc(HexNum(vPacket(2)(lChar), 1))
                            End If
                        Next
                    Case Else
                End Select
                
            End If
        End If
    Next
End Sub

Public Sub ShowPage(ByVal sPage As String, ByVal sSubPage As String)
    Dim lColumn As Long
    Dim lRow As Long
    Dim lPage As Long
    Dim lSubPage As Long
    Dim yPage(39, 24) As Byte
    
    
    lPage = ConvertBase(sPage, 16)
    lSubPage = ConvertBase(sSubPage, 16)
    
    'Debug.Print mbPageLoaded(lPage)
    For lRow = 0 To 31
        Debug.Print vbCrLf;
        For lColumn = 0 To 39
            If lRow < 25 Then
                yPage(lColumn, lRow) = myPages(lColumn, lRow, lSubPage, lPage)
            End If
            Debug.Print Chr$(myPages(lColumn, lRow, lSubPage, lPage));
        Next
    Next
    Open App.Path & "\output.mem" For Binary As #1
    Put #1, , yPage
    Close #1
End Sub

Public Function GetPage(ByVal lPage As Long, ByVal lSubPage As Long) As Byte()
    Dim lColumn As Long
    Dim lRow As Long
    Dim yPage() As Byte
    Dim yCode As Byte
    Dim bDouble As Boolean
    
    ReDim yPage(999)

    For lRow = 0 To 31
        bDouble = False
        For lColumn = 0 To 39
            If lRow < 25 Then
                yCode = myPages(lColumn, lRow, lSubPage, lPage)
                If yCode < 32 Then
                    yCode = yCode + 128
                End If
                yPage(lColumn + lRow * 40) = yCode
                If yCode = 141 Then
                    bDouble = True
                End If
            End If
        Next
        
        If bDouble Then
            If lRow < 24 Then
                For lColumn = 0 To 39
                    yPage(lColumn + (lRow + 1) * 40) = yPage(lColumn + lRow * 40)
                Next
                lRow = lRow + 1
            End If
        End If
    Next
    
    For lColumn = 0 To 7
        yPage(lColumn) = 32
    Next
    
    If lPage = 896 Then
        For lColumn = 0 To 2
            yPage(lColumn + 0 * 40) = yPage(lColumn + 5 * 40)
        Next
    End If
    
    GetPage = yPage
End Function

Private Function AppendMask(yPage() As Byte, yMask() As Byte)
    Dim lPageSize As Long
    
    lPageSize = UBound(yPage) + 1
    
    ReDim Preserve yPage(UBound(yMask) + 1 + lPageSize - 1)
    CopyMemory yPage(UBound(yMask) + 1), yPage(0), lPageSize
    CopyMemory yPage(0), yMask(0), UBound(yMask) + 1
End Function

Private Function RemoveBlankLines(yData() As Byte) As Byte()
    Dim lLine As Long
    Dim lColumn As Long
    Dim sBinary As String
    Dim lIsBlank As Long
    Dim yMask(3) As Byte
    Dim lCopyLineTo As Long
    
    For lLine = 0 To 999 Step 40
        lIsBlank = 1
        For lColumn = 0 To 39
            Select Case yData(lLine + lColumn)
                Case 33 To 127, 157
                    lIsBlank = 0
                    Exit For
            End Select
        Next
        sBinary = sBinary + CStr(lIsBlank)
        If lIsBlank = 0 Then
            For lColumn = 0 To 39
                yData(lCopyLineTo + lColumn) = yData(lLine + lColumn)
            Next
            lCopyLineTo = lCopyLineTo + 40
        End If
    Next
    sBinary = sBinary & "0000000"
    For lColumn = 1 To 32 Step 8
        yMask((lColumn - 1) \ 8) = ConvertBase(Mid$(sBinary, lColumn, 8), 2)
    Next
    If lCopyLineTo = 0 Then
        ReDim Preserve yData(7)
    Else
        ReDim Preserve yData(lCopyLineTo - 1)
    End If
    RemoveBlankLines = yMask
End Function

Private Sub CompressData(yData() As Byte)
    Dim lStep As Long
    Dim yCompressed(6) As Byte
    Dim lOffset As Long
    Dim lMultiplier As Long
    Dim lNewData As Long
    Dim lCopyPosition As Long
    
    For lStep = 0 To UBound(yData) - 1 Step 8
        lMultiplier = 2
        Erase yCompressed
        For lOffset = 0 To 7
            lNewData = CLng(yData(lStep + lOffset) And &H7F&) * lMultiplier
            If lOffset < 7 Then
                yCompressed(lOffset) = lNewData And &HFF&
            End If
            If lOffset > 0 Then
                yCompressed(lOffset - 1) = yCompressed(lOffset - 1) Or (lNewData \ 256)
            End If
            
            lMultiplier = lMultiplier * 2
        Next
        CopyMemory yData(lCopyPosition), yCompressed(0), 7
        lCopyPosition = lCopyPosition + 7
    Next
    ReDim Preserve yData(lCopyPosition - 1)
End Sub

Private Function Page(ByVal sPage As String) As Long
    Page = ConvertBase(sPage, 16)
End Function

Public Function ConvertBase(ByVal sNumber As String, ByVal lBase As Long) As Long
    Dim lIndex As Long
    Const sChars As String = "0123456789ABCDEF"
    Dim lValue As Long
    
    For lIndex = 1 To Len(sNumber)
        lValue = lValue * lBase
        lValue = lValue + InStr(sChars, Mid$(sNumber, lIndex, 1)) - 1
    Next
    ConvertBase = lValue
End Function

Public Function Pad(ByVal sNumber As String, ByVal lPlaces As Long) As String
    If Len(sNumber) >= lPlaces Then
        Pad = Right$(sNumber, lPlaces)
    Else
        Pad = String$(lPlaces - Len(sNumber), "0") & sNumber
    End If
End Function


Public Function HexNum(ByVal lNumber As Long, ByVal iPlaces As Integer) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
End Function


Private Function DecodePacket(ByVal lPos As Long) As Variant
    Dim lLineCode As Long
    Dim lMagazine As Long
    Dim lRow As Long
    Dim yData(39) As Byte
    Dim yHeader(7) As Byte
    Dim lIndex As Long
    
    Dim lValue1 As Long
    Dim lValue2 As Long
    
    lValue1 = InverseHamming(myFile(lPos))
    lValue2 = InverseHamming(myFile(lPos + 1))
    
    If lValue1 = -1 Or lValue2 = -1 Then
        DecodePacket = Array(-1, -1, yData())
        Exit Function
    End If
    
    lLineCode = lValue1 + lValue2 * 16
    lRow = (lLineCode And &HF8) \ 8
    lMagazine = lLineCode And &H7
    
    'Debug.Print vbCrLf & lMagazine & ":" & lRow;
    If lRow = 0 Then
        For lIndex = 0 To 7
            lValue1 = InverseHamming(myFile(lPos + 2 + lIndex))
            If lValue1 = -1 Then
                DecodePacket = Array(-1, -1, yData())
                Exit Function
            End If
            yData(lIndex) = InverseHamming(myFile(lPos + 2 + lIndex))
        Next
        For lIndex = 8 To 39
            yData(lIndex) = myFile(lPos + 2 + lIndex) And &H7F&
            'Debug.Print Chr$(yData(lIndex));
        Next
    Else
        For lIndex = 0 To 39
            yData(lIndex) = myFile(lPos + 2 + lIndex) And &H7F&
            'Debug.Print Chr$(yData(lIndex));
        Next
        
    End If
    DecodePacket = Array(lMagazine, lRow, yData)
End Function

Private Function InverseHamming(ByVal lValue As Long) As Long
    Dim lHammed(255) As Long
        
    lHammed(&H0) = &H1
    lHammed(&H1) = -1
    lHammed(&H2) = &H1
    lHammed(&H3) = &H1
    lHammed(&H4) = -1
    lHammed(&H5) = &H0
    lHammed(&H6) = &H1
    lHammed(&H7) = -1
    lHammed(&H8) = -1
    lHammed(&H9) = &H2
    lHammed(&HA) = &H1
    lHammed(&HB) = -1
    lHammed(&HC) = &HA
    lHammed(&HD) = -1
    lHammed(&HE) = -1
    lHammed(&HF) = &H7
    lHammed(&H10) = -1
    lHammed(&H11) = &H0
    lHammed(&H12) = &H1
    lHammed(&H13) = -1
    lHammed(&H14) = &H0
    lHammed(&H15) = &H0
    lHammed(&H16) = -1
    lHammed(&H17) = &H0
    lHammed(&H18) = &H6
    lHammed(&H19) = -1
    lHammed(&H1A) = -1
    lHammed(&H1B) = &HB
    lHammed(&H1C) = -1
    lHammed(&H1D) = &H0
    lHammed(&H1E) = &H3
    lHammed(&H1F) = -1
    lHammed(&H20) = -1
    lHammed(&H21) = &HC
    lHammed(&H22) = &H1
    lHammed(&H23) = -1
    lHammed(&H24) = &H4
    lHammed(&H25) = -1
    lHammed(&H26) = -1
    lHammed(&H27) = &H7
    lHammed(&H28) = &H6
    lHammed(&H29) = -1
    lHammed(&H2A) = -1
    lHammed(&H2B) = &H7
    lHammed(&H2C) = -1
    lHammed(&H2D) = &H7
    lHammed(&H2E) = &H7
    lHammed(&H2F) = &H7
    lHammed(&H30) = &H6
    lHammed(&H31) = -1
    lHammed(&H32) = -1
    lHammed(&H33) = &H5
    lHammed(&H34) = -1
    lHammed(&H35) = &H0
    lHammed(&H36) = &HD
    lHammed(&H37) = -1
    lHammed(&H38) = &H6
    lHammed(&H39) = &H6
    lHammed(&H3A) = &H6
    lHammed(&H3B) = -1
    lHammed(&H3C) = &H6
    lHammed(&H3D) = -1
    lHammed(&H3E) = -1
    lHammed(&H3F) = &H7
    lHammed(&H40) = -1
    lHammed(&H41) = &H2
    lHammed(&H42) = &H1
    lHammed(&H43) = -1
    lHammed(&H44) = &H4
    lHammed(&H45) = -1
    lHammed(&H46) = -1
    lHammed(&H47) = &H9
    lHammed(&H48) = &H2
    lHammed(&H49) = &H2
    lHammed(&H4A) = -1
    lHammed(&H4B) = &H2
    lHammed(&H4C) = -1
    lHammed(&H4D) = &H2
    lHammed(&H4E) = &H3
    lHammed(&H4F) = -1
    lHammed(&H50) = &H8
    lHammed(&H51) = -1
    lHammed(&H52) = -1
    lHammed(&H53) = &H5
    lHammed(&H54) = -1
    lHammed(&H55) = &H0
    lHammed(&H56) = &H3
    lHammed(&H57) = -1
    lHammed(&H58) = -1
    lHammed(&H59) = &H2
    lHammed(&H5A) = &H3
    lHammed(&H5B) = -1
    lHammed(&H5C) = &H3
    lHammed(&H5D) = -1
    lHammed(&H5E) = &H3
    lHammed(&H5F) = &H3
    lHammed(&H60) = &H4
    lHammed(&H61) = -1
    lHammed(&H62) = -1
    lHammed(&H63) = &H5
    lHammed(&H64) = &H4
    lHammed(&H65) = &H4
    lHammed(&H66) = &H4
    lHammed(&H67) = -1
    lHammed(&H68) = -1
    lHammed(&H69) = &H2
    lHammed(&H6A) = &HF
    lHammed(&H6B) = -1
    lHammed(&H6C) = &H4
    lHammed(&H6D) = -1
    lHammed(&H6E) = -1
    lHammed(&H6F) = &H7
    lHammed(&H70) = -1
    lHammed(&H71) = &H5
    lHammed(&H72) = &H5
    lHammed(&H73) = &H5
    lHammed(&H74) = &H4
    lHammed(&H75) = -1
    lHammed(&H76) = -1
    lHammed(&H77) = &H5
    lHammed(&H78) = &H6
    lHammed(&H79) = -1
    lHammed(&H7A) = -1
    lHammed(&H7B) = &H5
    lHammed(&H7C) = -1
    lHammed(&H7D) = &HE
    lHammed(&H7E) = &H3
    lHammed(&H7F) = -1
    lHammed(&H80) = -1
    lHammed(&H81) = &HC
    lHammed(&H82) = &H1
    lHammed(&H83) = -1
    lHammed(&H84) = &HA
    lHammed(&H85) = -1
    lHammed(&H86) = -1
    lHammed(&H87) = &H9
    lHammed(&H88) = &HA
    lHammed(&H89) = -1
    lHammed(&H8A) = -1
    lHammed(&H8B) = &HB
    lHammed(&H8C) = &HA
    lHammed(&H8D) = &HA
    lHammed(&H8E) = &HA
    lHammed(&H8F) = -1
    lHammed(&H90) = &H8
    lHammed(&H91) = -1
    lHammed(&H92) = -1
    lHammed(&H93) = &HB
    lHammed(&H94) = -1
    lHammed(&H95) = &H0
    lHammed(&H96) = &HD
    lHammed(&H97) = -1
    lHammed(&H98) = -1
    lHammed(&H99) = &HB
    lHammed(&H9A) = &HB
    lHammed(&H9B) = &HB
    lHammed(&H9C) = &HA
    lHammed(&H9D) = -1
    lHammed(&H9E) = -1
    lHammed(&H9F) = &HB
    lHammed(&HA0) = &HC
    lHammed(&HA1) = &HC
    lHammed(&HA2) = -1
    lHammed(&HA3) = &HC
    lHammed(&HA4) = -1
    lHammed(&HA5) = &HC
    lHammed(&HA6) = &HD
    lHammed(&HA7) = -1
    lHammed(&HA8) = -1
    lHammed(&HA9) = &HC
    lHammed(&HAA) = &HF
    lHammed(&HAB) = -1
    lHammed(&HAC) = &HA
    lHammed(&HAD) = -1
    lHammed(&HAE) = -1
    lHammed(&HAF) = &H7
    lHammed(&HB0) = -1
    lHammed(&HB1) = &HC
    lHammed(&HB2) = &HD
    lHammed(&HB3) = -1
    lHammed(&HB4) = &HD
    lHammed(&HB5) = -1
    lHammed(&HB6) = &HD
    lHammed(&HB7) = &HD
    lHammed(&HB8) = &H6
    lHammed(&HB9) = -1
    lHammed(&HBA) = -1
    lHammed(&HBB) = &HB
    lHammed(&HBC) = -1
    lHammed(&HBD) = &HE
    lHammed(&HBE) = &HD
    lHammed(&HBF) = -1
    lHammed(&HC0) = &H8
    lHammed(&HC1) = -1
    lHammed(&HC2) = -1
    lHammed(&HC3) = &H9
    lHammed(&HC4) = -1
    lHammed(&HC5) = &H9
    lHammed(&HC6) = &H9
    lHammed(&HC7) = &H9
    lHammed(&HC8) = -1
    lHammed(&HC9) = &H2
    lHammed(&HCA) = &HF
    lHammed(&HCB) = -1
    lHammed(&HCC) = &HA
    lHammed(&HCD) = -1
    lHammed(&HCE) = -1
    lHammed(&HCF) = &H9
    lHammed(&HD0) = &H8
    lHammed(&HD1) = &H8
    lHammed(&HD2) = &H8
    lHammed(&HD3) = -1
    lHammed(&HD4) = &H8
    lHammed(&HD5) = -1
    lHammed(&HD6) = -1
    lHammed(&HD7) = &H9
    lHammed(&HD8) = &H8
    lHammed(&HD9) = -1
    lHammed(&HDA) = -1
    lHammed(&HDB) = &HB
    lHammed(&HDC) = -1
    lHammed(&HDD) = &HE
    lHammed(&HDE) = &H3
    lHammed(&HDF) = -1
    lHammed(&HE0) = -1
    lHammed(&HE1) = &HC
    lHammed(&HE2) = &HF
    lHammed(&HE3) = -1
    lHammed(&HE4) = &H4
    lHammed(&HE5) = -1
    lHammed(&HE6) = -1
    lHammed(&HE7) = &H9
    lHammed(&HE8) = &HF
    lHammed(&HE9) = -1
    lHammed(&HEA) = &HF
    lHammed(&HEB) = &HF
    lHammed(&HEC) = -1
    lHammed(&HED) = &HE
    lHammed(&HEE) = &HF
    lHammed(&HEF) = -1
    lHammed(&HF0) = &H8
    lHammed(&HF1) = -1
    lHammed(&HF2) = -1
    lHammed(&HF3) = &H5
    lHammed(&HF4) = -1
    lHammed(&HF5) = &HE
    lHammed(&HF6) = &HD
    lHammed(&HF7) = -1
    lHammed(&HF8) = -1
    lHammed(&HF9) = &HE
    lHammed(&HFA) = &HF
    lHammed(&HFB) = -1
    lHammed(&HFC) = &HE
    lHammed(&HFD) = &HE
    lHammed(&HFE) = -1
    lHammed(0) = &HE
    
    InverseHamming = lHammed(lValue)
End Function

Private Function InverseHamming2418(ByVal lValue As Long) As Long
    Dim sBinary As String
    Dim lMultiplier As Long
    Dim lBitPos As Long
    Dim sResult As String
    Dim lCheck(5) As Long
    Dim lStepSize As Long
    Dim lCheckIndex As Long
    Dim lStepPos As Long
    Dim lRunPos As Long
    
    lMultiplier = 1
    sBinary = Pad(ConvertBase(lValue, 2), 24)
    
    For lBitPos = 1 To 24
        If lBitPos <> lMultiplier Then
            sResult = Mid$(sBinary, 24 - lBitPos + 1, 1) & sResult
        Else
            lMultiplier = lMultiplier * 2
        End If
    Next
    
    For lBitPos = 0 To 23
        lCheck(0) = lCheck(0) Xor Val(Mid$(sBinary, 24 - lBitPos, 1))
    Next
    
    lStepSize = 2

    For lCheckIndex = 1 To 5
        For lStepPos = lStepSize \ 2 - 1 To 23 Step lStepSize
            For lRunPos = 0 To lStepSize \ 2 - 1
                lCheck(lCheckIndex) = lCheck(lCheckIndex) Xor Val(Mid$(sBinary, 24 - (lStepPos + lRunPos), 1))
            Next
        Next
        lStepSize = lStepSize * 2
    Next
    
    InverseHamming2418 = ConvertBase(sResult, 2)
End Function
