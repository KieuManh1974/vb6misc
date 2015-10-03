Attribute VB_Name = "Algorithm"
Option Explicit

' // vcrplus.h
Private Const INVALID = -987
Private Const LIMIT = 2202
Private Const SOFTLIMIT = 2000

' // scramble.h
Private Const EncodeKey As String = "68150631" ' Key 2,8
Private Const DecodeKey As String = "9371" ' Key 16,4

' //  Decode.c
Private MonthArray As Variant

Private Type s_tab
    start As Integer
    leng As Integer
End Type

Private SlotTable(LIMIT) As Variant ' Slot times and durations
Private TList As Variant
Private DList As Variant

' // b3.h
Private Sub Decode_Right3Digits(ByVal Right3Digits As Integer, ByRef day As Integer, ByRef Right5BinaryDigits As Integer)
    day = ((Right3Digits - 1) / 32) + 1 ' Top bits
    Right5BinaryDigits = (Right3Digits - 1) Mod 32 ' Bottom 5 bits
End Sub

Private Function Encode_Right3Digits(ByVal day As Integer, ByVal Right5BinaryDigits As Integer)
    Encode_Right3Digits = Right5BinaryDigits + (32 * (day - 1)) + 1
End Function

Private Function DecodeRight3Digits2(sRight3Digits As String) As String
    DecodeRight3Digits2 = Convert(sRight3Digits, 10, 32)
End Function

' // Will produce a binary string of a decimal number
Public Function Binary2(sNumber As String, iSize As Integer) As String
    Binary2 = Pad(Convert(sNumber, 10, 2), iSize)
End Function

' // Will produce a decimal number of a binary string
Public Function Dec2(sNumber As String) As String
    Dec2 = Convert(sNumber, 2, 10)
End Function

' // Interleaves bits
Private Sub Interleave(SlotIndex As Long, cval As Long, t8c5 As Long, t2c1 As Long)
    Dim tblidxBinary As String
    Dim cvalBinary As String
    
    tblidxBinary = Binary2(CStr(SlotIndex), 11)
    cvalBinary = Binary2(CStr(cval), 4)
    
    t8c5 = Dec2(Left(tblidxBinary, 7) & Mid(cvalBinary, 4 - 3, 1) & Mid(cvalBinary, 4 - 2, 1) & Mid(tblidxBinary, 11 - 3, 1))
    t2c1 = Dec2(Mid(tblidxBinary, 11 - 2, 1) & Mid(cvalBinary, 4 - 1, 1) & Mid(tblidxBinary, 11 - 1, 1) & Mid(cvalBinary, 4 - 0, 1) & Mid(tblidxBinary, 11 - 0, 1))
End Sub

Private Sub Deinterleave(ByVal t8c5 As Long, ByVal t2c1 As Integer, SlotIndex As Long, cval As Integer)
    Dim t8c5binary As String
    Dim t2c1binary As String
    
    t8c5binary = Binary2(CStr(t8c5), 10)
    t2c1binary = Binary2(CStr(t2c1), 10)

    SlotIndex = Dec2(Left(t8c5binary, 7) & Right(t8c5binary, 1) & Mid(t2c1binary, 10 - 4, 1) & Mid(t2c1binary, 10 - 2, 1) & Mid(t2c1binary, 10 - 0, 1))
    cval = Dec2(Mid(t8c5binary, 10 - 2, 1) & Mid(t8c5binary, 10 - 1, 1) & Mid(t2c1binary, 10 - 3, 1) & Mid(t2c1binary, 10 - 1, 1))

End Sub

Private Sub Deinterleave2(ByVal A As Long, ByVal B As Integer, iSlotIndex As Long, iChannel As Integer)
    Dim Abin As String
    Dim Bbin As String
    
    Abin = Binary2(CStr(A), 10)
    Bbin = Binary2(CStr(B), 5)

    iSlotIndex = Dec2(Slice(Abin, 1, 7) & Slice(Abin, 10) & Slice(Bbin, 1) & Slice(Bbin, 3) & Slice(Bbin, 5))
    iChannel = Dec2(Slice(Abin, 8, 9) & Slice(Bbin, 2) & Slice(Bbin, 4))
    
    Debug.Print Abin & "  " & Bbin
    Debug.Print iSlotIndex & " " & iChannel
End Sub

' // misc.h
Private Function EndTime(ByVal start As Integer, ByVal dur As Integer) As Integer
    Dim min As Integer
    Dim hr As Integer
    
    min = (start Mod 100) + dur
    
    hr = min \ 60
    min = min Mod 60
    hr = (hr + start \ 100) Mod 24
    EndTime = hr * 100 + min
End Function

' // scramble.h
Public Function CrossMultiply2(ByVal sValue As String, ByVal sKey As String) As String
    Dim iValueLen As Integer
    iValueLen = Len(StripZero(sValue))
    Do
        CrossMultiply2 = Pad(Pad(Multiply(sValue, sKey, 10, 0), iValueLen), 8)
        sValue = CrossMultiply2
    Loop Until Mid$(CrossMultiply2, Len(CrossMultiply2) - iValueLen + 1, 1) <> "0"
End Function

Private Function UnmapTop(ByVal day As Integer, ByVal year As Integer, ByVal top As Long, ByVal digits As Integer) As Long
    Dim d2 As Long
    Dim d1 As Long
    Dim d0 As Long
    Dim y As Long
    Dim poot As Long
    Dim n2 As Long
    Dim n1 As Long
    Dim n0 As Long
    Dim f3 As Long
    Dim f2 As Long
    Dim f1 As Long
    Dim f0 As Long
    Dim p3 As Long
    Dim p2 As Long
    Dim p1 As Long
    
    d2 = top \ 100
    d1 = (top Mod 100) \ 10
    d0 = top Mod 10
    
    ' / generate key (P3P2P1F0) and reverse key (F3F2F1F0)
    f0 = 1
    y = year Mod 16
    p1 = (y + 1) Mod 10
    f1 = 10 - p1
    
    p2 = (((y + 1) * (y + 2)) \ 2) Mod 10
    f2 = 10 - ((p2 + f1 * p1) Mod 10)
    
    p3 = (((y + 1) * (y + 2) * (y + 3)) \ 6) Mod 10
    f3 = 10 - ((p3 + f1 * p2 + f2 * p1) Mod 10)
    
    If digits = 1 Then
        n0 = (d0 * f0 + day * f1) Mod 10
        n1 = 0
        n2 = 0
    End If
    
    If digits = 2 Then
        n0 = (d0 * f0 + d1 * f1 + day * f2) Mod 10
        n1 = (d1 * f0 + day * f1) Mod 10
        n2 = 0
    End If
    
    If digits = 3 Then
        n0 = (d0 * f0 + d1 * f1 + d2 * f2 + day * f3) Mod 10
        n1 = (d1 * f0 + d2 * f1 + day * f2) Mod 10
        n2 = (d2 * f0 + day * f1) Mod 10
    End If
    
    poot = 100 * n2 + 10 * n1 + n0
    
    UnmapTop = poot
End Function

' // vcrplus.h -
' // qlookup.c

Private Function FindSlotIndex(StartTime As Integer, Duration As Integer) As Long
    Dim j As Long
    
    For j = 0 To SOFTLIMIT - 1
        If SlotTable(j)(0) = StartTime And SlotTable(j)(1) = Duration Then
            FindSlotIndex = j
            Exit Function
        End If
    Next
End Function

Private Function ScanForStart(start As Integer, prev As Long) As Long
    Dim j As Long
    
    CreateSlotTable
    
    For j = prev + 1 To SOFTLIMIT - 1
        If SlotTable(j).start = start Then
            ScanForStart = j
        End If
    Next
End Function

Private Sub Lookup(ByVal i As Long, StartTime As Integer, Duration As Integer)
    If i > LIMIT Then
        'Debug.Print "Illegal table index"
        Exit Sub
    End If
    
    If i > SOFTLIMIT Then
        StartTime = INVALID
        Duration = INVALID
        Exit Sub
    End If
    
'    Open App.Path & "\slots.txt" For Output As #1
'    For i = 0 To 2202
'        Print #1, CStr(SlotTable(i)(0)) & "," & CStr(SlotTable(i)(1))
'    Next
'    Close #1
    
    StartTime = SlotTable(i)(0)
    Duration = SlotTable(i)(1)
End Sub

Private Function GetDuration(Index As Long) As Integer
    GetDuration = SlotTable(Index).leng
End Function

Private Sub CreateSlotTable()
    Dim i As Integer
    
    '/* insert movie (and other) structured tables */
    TList = Array(5, 10, 20, 25, _
            0, 5, 10, 15, 20, 25, _
            0, 5, 10, 15, 20, 25, _
            0, 5, 10, 15, 20, 25, _
            0, 5, 10, 15, 20, 25, _
            0, 5, 10, 15, 20, 25)
            
    DList = Array(0, 0, 0, 0, _
            -5, -5, -5, -5, -5, -5, _
            -10, -10, -10, -10, -10, -10, _
            -15, -15, -15, -15, -15, -15, _
            -20, -20, -20, -20, -20, -20, _
            -25, -25, -25, -25, -25, -25)
            
    Dim temp As New Collection
    '/* first 192 entries are erratic, so pre-load them */

    temp.Add Array(1830, 30)
    temp.Add Array(1600, 30)
    temp.Add Array(1930, 30)
    temp.Add Array(1630, 30)
    temp.Add Array(1530, 30)
    temp.Add Array(1730, 30)
    temp.Add Array(1800, 30)
    temp.Add Array(1430, 30)
    temp.Add Array(1900, 30)
    temp.Add Array(1700, 60)
    temp.Add Array(1400, 30)
    temp.Add Array(2030, 30)
    temp.Add Array(1700, 30)
    temp.Add Array(1600, 120)
    temp.Add Array(2000, 30)
    temp.Add Array(1500, 30)
    temp.Add Array(2000, 120)
    temp.Add Array(2100, 120)
    temp.Add Array(2000, 60)
    temp.Add Array(1800, 120)
    temp.Add Array(1900, 60)
    temp.Add Array(2200, 60)
    temp.Add Array(2100, 60)
    temp.Add Array(1400, 120)
    temp.Add Array(1500, 60)
    temp.Add Array(2200, 120)
    temp.Add Array(1130, 30)
    temp.Add Array(1100, 30)
    temp.Add Array(2300, 30)
    temp.Add Array(1600, 60)
    temp.Add Array(2100, 90)
    temp.Add Array(2100, 30)
    temp.Add Array(1230, 30)
    temp.Add Array(1330, 30)
    temp.Add Array(930, 30)
    temp.Add Array(1300, 60)
    temp.Add Array(2130, 30)
    temp.Add Array(1200, 60)
    temp.Add Array(1000, 120)
    temp.Add Array(1800, 60)
    temp.Add Array(2200, 30)
    temp.Add Array(1200, 30)
    temp.Add Array(800, 30)
    temp.Add Array(830, 30)
    temp.Add Array(1700, 120)
    temp.Add Array(900, 30)
    temp.Add Array(2230, 30)
    temp.Add Array(1030, 30)
    temp.Add Array(1900, 120)
    temp.Add Array(730, 30)
    temp.Add Array(2300, 60)
    temp.Add Array(1000, 60)
    temp.Add Array(700, 30)
    temp.Add Array(1300, 30)
    temp.Add Array(700, 120)
    temp.Add Array(1100, 60)
    temp.Add Array(1400, 60)
    temp.Add Array(1000, 30)
    temp.Add Array(800, 120)
    temp.Add Array(2330, 30)
    temp.Add Array(1300, 120)
    temp.Add Array(1200, 120)
    temp.Add Array(900, 120)
    temp.Add Array(630, 30)
    temp.Add Array(1800, 90)
    temp.Add Array(600, 30)
    temp.Add Array(530, 30)
    temp.Add Array(0, 30)
    temp.Add Array(2330, 120)
    temp.Add Array(2200, 90)
    temp.Add Array(1300, 90)
    temp.Add Array(900, 60)
    temp.Add Array(1630, 90)
    temp.Add Array(1600, 90)
    temp.Add Array(1430, 90)
    temp.Add Array(2000, 90)
    temp.Add Array(1830, 90)
    temp.Add Array(600, 60)
    temp.Add Array(1200, 90)
    temp.Add Array(30, 30)
    temp.Add Array(130, 120)
    temp.Add Array(0, 60)
    temp.Add Array(1700, 90)
    temp.Add Array(0, 120)
    temp.Add Array(800, 60)
    temp.Add Array(700, 60)
    temp.Add Array(2130, 120)
    temp.Add Array(500, 30)
    temp.Add Array(1530, 90)
    temp.Add Array(1130, 120)
    temp.Add Array(1100, 120)
    temp.Add Array(830, 90)
    temp.Add Array(2230, 90)
    temp.Add Array(900, 90)
    temp.Add Array(2130, 90)
    temp.Add Array(1630, 120)
    temp.Add Array(2330, 60)
    temp.Add Array(100, 120)
    temp.Add Array(1400, 90)
    temp.Add Array(130, 30)
    temp.Add Array(330, 120)
    temp.Add Array(1500, 90)
    temp.Add Array(1500, 120)
    temp.Add Array(2300, 120)
    temp.Add Array(1900, 90)
    temp.Add Array(800, 90)
    temp.Add Array(430, 30)
    temp.Add Array(300, 30)
    temp.Add Array(1330, 120)
    temp.Add Array(1000, 90)
    temp.Add Array(700, 90)
    temp.Add Array(100, 30)
    temp.Add Array(2330, 90)
    temp.Add Array(330, 30)
    temp.Add Array(200, 30)
    temp.Add Array(2230, 120)
    temp.Add Array(400, 30)
    temp.Add Array(600, 120)
    temp.Add Array(400, 120)
    temp.Add Array(230, 30)
    temp.Add Array(630, 90)
    temp.Add Array(30, 60)
    temp.Add Array(2230, 60)
    temp.Add Array(100, 60)
    temp.Add Array(30, 120)
    temp.Add Array(2300, 90)
    temp.Add Array(1630, 60)
    temp.Add Array(830, 60)
    temp.Add Array(0, 90)
    temp.Add Array(1930, 120)
    temp.Add Array(930, 120)
    temp.Add Array(2030, 90)
    temp.Add Array(500, 60)
    temp.Add Array(1730, 60)
    temp.Add Array(200, 120)
    temp.Add Array(1930, 90)
    temp.Add Array(930, 90)
    temp.Add Array(1730, 120)
    temp.Add Array(630, 120)
    temp.Add Array(1830, 60)
    temp.Add Array(1430, 60)
    temp.Add Array(1130, 90)
    temp.Add Array(30, 90)
    temp.Add Array(830, 120)
    temp.Add Array(1030, 90)
    temp.Add Array(1430, 120)
    temp.Add Array(100, 90)
    temp.Add Array(730, 120)
    temp.Add Array(2030, 120)
    temp.Add Array(300, 90)
    temp.Add Array(300, 120)
    temp.Add Array(1330, 90)
    temp.Add Array(1230, 90)
    temp.Add Array(230, 90)
    temp.Add Array(2130, 60)
    temp.Add Array(1130, 60)
    temp.Add Array(1830, 120)
    temp.Add Array(630, 60)
    temp.Add Array(530, 60)
    temp.Add Array(200, 60)
    temp.Add Array(1530, 120)
    temp.Add Array(730, 60)
    temp.Add Array(600, 90)
    temp.Add Array(1730, 90)
    temp.Add Array(400, 60)
    temp.Add Array(730, 90)
    temp.Add Array(430, 90)
    temp.Add Array(430, 60)
    temp.Add Array(130, 90)
    temp.Add Array(1230, 120)
    temp.Add Array(130, 60)
    temp.Add Array(230, 120)
    temp.Add Array(1930, 60)
    temp.Add Array(300, 60)
    temp.Add Array(1030, 120)
    temp.Add Array(200, 90)
    temp.Add Array(330, 60)
    temp.Add Array(500, 120)
    temp.Add Array(930, 60)
    temp.Add Array(230, 60)
    temp.Add Array(2030, 60)
    temp.Add Array(400, 90)
    temp.Add Array(1530, 60)
    temp.Add Array(430, 120)
    temp.Add Array(1330, 60)
    temp.Add Array(1230, 60)
    temp.Add Array(330, 90)
    temp.Add Array(1030, 60)
    temp.Add Array(500, 90)
    temp.Add Array(530, 120)
    temp.Add Array(530, 90)
    temp.Add Array(1100, 90)
    
'    Dim vItem As Variant
'    Dim oFS As New FileSystemObject
'    Dim oTS As TextStream
'
'    Set oTS = oFS.CreateTextFile(App.Path & "\slots.txt", True)
'
'    For Each vItem In temp
'        oTS.WriteLine 2 * vItem(0) \ 100 + (vItem(0) Mod 100) \ 30 + 48 * (vItem(1) - 30) \ 30
'    Next
'    oTS.Close
    
    For i = 0 To LIMIT
        SlotTable(i) = Array(INVALID, INVALID)
    Next
    
    For i = 0 To 191
        SlotTable(i) = Array(temp(i + 1)(0), temp(i + 1)(1))
    Next
    
    FillQHB 192, 30    ' 30min progs starting 0015 -> 2345 */
    FillQHB 240, 60
    FillQHB 288, 90
    FillQHB 336, 120
    Fill 384, 2230, 30     ' 30 ->5 min progs starting 2230 -> 2255 */
    Fill 418, 2300, 30
    Fill 452, 2330, 30
    Fill 486, 1930, 90     ' 90 -> 65min progs starting 1930 -> 1955 */
    Fill 520, 2300, 90     ' 90 -> 65min progs starting 2300 -> 2325 */
    Fill 554, 2330, 90
    Fill 588, 2130, 120    ' 120 -> 95min progs starting 2130 -> 2155 */
    Fill 622, 2200, 120
    Fill 656, 2300, 120    ' 120 -> 95min progs starting 2300 -> 2325 */
    Fill 690, 2330, 120
    Fill 724, 0, 120
    Fill 758, 30, 120
    Fill 792, 100, 120
    Fill 826, 130, 120
    Fill 860, 1730, 60     ' 60 -> 35min progs starting 1730 -> 1755 */
    Fill 894, 1800, 60
    Fill 928, 1830, 60
    Fill 962, 1900, 60
    Fill 996, 1930, 60
    Fill 1030, 2000, 60
    Fill 1064, 2030, 60
    Fill 1098, 200, 120    ' 120 -> 95min progs starting 0200 -> 0225 */
    Fill 1132, 230, 120
    Fill 1166, 300, 120
    Fill 1200, 330, 120
    Fill 1234, 400, 120
    Fill 1268, 1200, 120
    Fill 1302, 1400, 120
    Fill 1336, 1530, 120
    Fill 1370, 2100, 60    ' 60 -> 35min progs starting 2100 -> 2125 */
    Fill 1404, 2130, 60
    Fill 1438, 2200, 60
    Fill 1472, 2230, 60
    Fill 1506, 2300, 60
    Fill 1540, 2330, 60
    Fill 1574, 1730, 30    ' 30 -> 5min progs starting 1730 -> 1755 */
    Fill 1608, 1800, 30
    Fill 1642, 1830, 30
    Fill 1676, 1900, 30
    Fill 1710, 1930, 30
    Fill 1744, 2000, 30
    Fill 1778, 2030, 30
    Fill 1812, 2100, 30
    Fill 1846, 2130, 30
    Fill 1880, 2200, 30
    FillHHB 1914, 150  ' 150min progs starting 2330 -> 0000 */
    FillHHB 1962, 180  ' 180min progs starting 2330 -> 0000 */
    
End Sub


Private Sub Fill(Index As Integer, time As Integer, dur As Integer)
    Dim i As Integer
    
    For i = 0 To 33
        SlotTable(i + Index) = Array(time + TList(i), dur + DList(i))
    Next
    
End Sub


Private Sub FillQHB(Index As Integer, dur As Integer)
    Dim i As Integer
    
    For i = 0 To 23
        SlotTable(Index + i * 2) = Array(i * 100 + 15, dur)
        SlotTable(Index + i * 2 + 1) = Array(i * 100 + 45, dur)
    Next
End Sub


Private Function FillHHB(Index As Integer, dur As Integer)
    Dim i As Integer
    
    For i = 23 To 0 Step -1
        SlotTable(Index + (23 - i) * 2) = Array(i * 100 + 30, dur)
        SlotTable(Index + (23 - i) * 2 + 1) = Array(i * 100, dur)
    Next
End Function


' //  Decode.c
Public Function Decode(ByVal ThisMonth As Integer, ByVal ThisDate As Integer, ByVal ThisYear As Integer, ByVal number As Long) As Variant
    Dim EncodedNumber As String
    Dim Left5Digits As String
    Dim Right3Digits As Integer
    Dim LeftBinaryDigits As Integer
    Dim Right5BinaryDigits As Integer
    Dim s5_out As Integer
    Dim ofout As Integer
    Dim mtout As Long
    Dim SlotIndex As Long
    Dim DayOut As Integer
    Dim Channel As Integer
    Dim StartTime As Integer
    Dim Duration As Integer
    
    MonthArray = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    CreateSlotTable
    
    ThisYear = ThisYear Mod 100
    
    If ThisMonth > 12 Or ThisMonth < 1 Or ThisDate < 1 Or ThisDate > 31 Then
        'Debug.Print "Invalid date"
        Exit Function
    End If
    
    If number < 1 Or number > 99999999 Then
        'Debug.Print "Sorry, plus code too long"
        Exit Function
    End If
    
    ofout = INVALID
    mtout = INVALID
    
    EncodedNumber = CrossMultiply2(CStr(number), GenerateKey(2, 8))
        
    Right3Digits = CInt(Right$(EncodedNumber, 3))
    Left5Digits = Left$(EncodedNumber, 5)
    LeftBinaryDigits = (Right3Digits - 1) \ 32
    Right5BinaryDigits = (Right3Digits - 1) Mod 32
    DayOut = LeftBinaryDigits + 1
        
    If DayOut < ThisDate Then
        ThisMonth = ThisMonth + 1
        If ThisMonth > 12 Then
            ThisMonth = 1
            ThisYear = (ThisYear + 1) Mod 100
        End If
    End If
    
    If number >= 1000 Then
        Offset DayOut, ThisYear, Left5Digits, ofout, mtout
    Else
        mtout = 0
        ofout = 0
    End If
        
    s5_out = (Right5BinaryDigits + (DayOut * (ThisMonth + 1)) + ofout) Mod 32

    Deinterleave2 mtout, s5_out, SlotIndex, Channel
        
    Channel = Channel + 1
    Lookup SlotIndex, StartTime, Duration
    
    Decode = Array(DayOut, ThisMonth, ThisYear, Channel, StartTime, Duration, EndTime(StartTime, Duration))
End Function

Private Sub Offset(ByVal iDay As Integer, ByVal iYear As Integer, ByVal sTop5Digits As String, OffsetOut As Integer, TopOut As Long)
    Dim i As Integer
    Dim Offset As Integer
    Dim iDigitCount As Long
    Dim d As String
    Dim MapTopX As String
    Dim sTop5Short As String
    
    iDigitCount = Len(CStr(Val(sTop5Digits)))
    sTop5Short = Right$(sTop5Digits, iDigitCount)
    For i = 1 To iDigitCount
        Offset = Offset + Val(Slice(sTop5Short, i))
    Next
    
    Do
        For i = 0 To (iYear Mod 16)
            d = CStr(iDay Mod 10) & sTop5Short
            MapTopX = Slice(Multiply(Reverse(GenerateKey(i, 8)), d, 10, 0), 2, 2 + iDigitCount - 1)
            Offset = Offset + Val(Right$(MapTopX, 1))
        Next
        sTop5Short = MapTopX
    Loop Until Slice(sTop5Short, 1) <> "0" Or Val(sTop5Short) = 0
    
    OffsetOut = Offset Mod 32
    TopOut = Val(sTop5Short)
End Sub

' // Encode.c
Public Function Encode(day As Integer, month As Integer, year As Integer, Channel As Integer, StartTime As Integer, Duration As Integer) As String
    Dim j As Integer
    Dim SlotIndex As Long
    Dim limit_ As Long
    Dim doneflag As Long ' FULLSEARCH
    Dim s5_out As Long
    Dim Right3Digits As Integer
    Dim Right5BinaryDigits As Integer
    Dim ofout As Integer
    Dim EncodedNumber As Long
    Dim Left5Digits As Long
    Dim number As Long
    Dim s4_out As Long
    
    MonthArray = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    CreateSlotTable
    
    year = year Mod 100
    
    SlotIndex = FindSlotIndex(StartTime, Duration)
    If SlotIndex = -1 Then
        Encode = -1 ' error
        Exit Function
    End If
    
    ' From them infer what must have been step 4 & step 5 results */
    Interleave SlotIndex, Channel - 1, s4_out, s5_out
    
    ' If the mapped_top is zero then top and offset are zero */
    If s4_out = 0 Then
        Left5Digits = 0
        ofout = 0
    Else
        Dim i As Integer
        Dim tmp As Long
        
        j = Len(CStr(s4_out))
        limit_ = 10 ^ j
        If j > 3 Then
            Encode = 0 ' needs higher digit coding
        End If
        
        limit_ = limit_ \ 10
        ofout = 0
        Left5Digits = s4_out
        ' Get a Left5Digits with same no digits as s4_out
        ' May have to loop several times
        Do
            ' Reverse the MapTop encryption
            Left5Digits = UnmapTop(day, year, Left5Digits, j)
            For i = 0 To (year Mod 16)
                'ofout = ofout + (MapTop(day, i, Left5Digits, j) Mod 10)
            Next
        Loop While Left5Digits < limit_
        
        ' Add sum of final Left5Digits's digits to offset
        tmp = Left5Digits
        
        While tmp > 0
            ofout = ofout + (tmp Mod 10)
            tmp = tmp \ 10
        Wend
        
        ofout = ofout Mod 32
        
    End If
    
    ' Have two of the three inputs to step 5; determine the rem
    For Right5BinaryDigits = 0 To 31
        j = (Right5BinaryDigits + (day * (month + 1)) + ofout) Mod 32
        If j = s5_out Then
            Exit For
        End If
    Next
    
    ' Assemble the output of step 1
    Right3Digits = Encode_Right3Digits(day, Right5BinaryDigits)
    EncodedNumber = Right3Digits + (1000 * Left5Digits)

    ' Invert the mixing
    'number = CrossMultiply(EncodedNumber, DecodeKey)
    number = CrossMultiply2(CStr(EncodedNumber), CStr(DecodeKey))
  
    Encode = number
    
End Function

Public Sub Test()
    Dim vDecode As Variant
    
    vDecode = Decode(3, 2, 2005, 10)
    Debug.Print vDecode(0)
    Debug.Print vDecode(1)
    Debug.Print vDecode(2)
    Debug.Print vDecode(3)
    Debug.Print vDecode(4)
    Debug.Print vDecode(5)
End Sub
