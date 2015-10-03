Attribute VB_Name = "Module1"
Option Explicit

Private Const g_iflag As Boolean = False

Private vTtbl As Variant
Private vOverSixTable As Variant

Public Sub Initialise()
    Dim oFS As New FileSystemObject
    Dim oTS As TextStream
    Dim sValue As String
    
    vTtbl = Array()
    
    Set oTS = oFS.OpenTextFile(App.Path & "\slots.txt")
    
    While Not oTS.AtEndOfStream
        sValue = oTS.ReadLine
        ReDim Preserve vTtbl(UBound(vTtbl) + 1) As Variant
        vTtbl(UBound(vTtbl)) = Val(sValue)
    Wend
    oTS.Close
    
    vOverSixTable = Array(&H25, &H13, &H34, &H14, &H52, &H25, &H25, &H43, _
    &H45, &H14, &H43, &H51, &H21, &H12, &H23, &H43, _
    &H51, &H32, &H35, &H52, &H34, &H25, &H34, &H52, _
    &H13, &H15, &H41, &H35, &H52, &H51, &H12)
End Sub

Public Function Bit(sNumber As String, lPos As Long) As String
    Bit = Mid$(sNumber, Len(sNumber) - lPos, 1)
End Function

Public Function SetBit(sNumber As String, lPos As Long, sValue As String)
    Mid$(sNumber, Len(sNumber) - lPos, 1) = sValue
End Function

Public Function Pad(sNumber As String, lLength As Long) As String
    If lLength - Len(sNumber) >= 0 Then
        Pad = String$(lLength - Len(sNumber), "0") & sNumber
    Else
        Pad = Right$(sNumber, lLength)
    End If
End Function

Public Function Convert(sNumber As String, lFromBase As Long, lToBase As Long) As String
    Dim lPos As Long
    Dim lNumber As Long
    Dim lMultiplier As Long
    
    For lPos = 1 To Len(sNumber)
        lNumber = lNumber * lFromBase
        lNumber = lNumber + Val(Mid$(sNumber, lPos, 1))
    Next
    
    Do
        Convert = CStr((lNumber Mod lToBase)) & Convert
        lNumber = lNumber \ lToBase
    Loop Until lNumber = 0
End Function

Public Function AddDigits(sNumber As String) As Long
    Dim lIndex As Long
    
    For lIndex = 1 To Len(sNumber)
        AddDigits = AddDigits + Val(Mid$(sNumber, lIndex, 1))
    Next
End Function

Public Sub twiddle_tt(tidx As Long, tp As Long, dp As Long, x1p As Long, x2p As Long)
    Dim tt As Long
    Dim x1 As Long
    Dim x2 As Long
    Dim t As Long
    Dim d As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim b As Long
    
    x1 = 0
    x2 = 0
    t1 = 0
    t2 = 0
    d = 0
    b = 0
    t = 0
    tt = tidx
    
    If tt >= 768 And tt <= 3647 Then
        t1 = ((tt - 768) Mod 10) + 1
        t2 = t1 * 5
        If t1 >= 6 Then
            x2 = t2 - 25
        Else
            x1 = t2
        End If
        tt = tt - 768
        tt = tt \ 10
    ElseIf tt > 3648 And tt <= 6527 Then
        t1 = ((tt - 3648) Mod 6) + 1
        If t1 = 1 Then
            x1 = 15
        Else
            x2 = (5 * t1) - 5
        End If
    ElseIf tt > 6528 And tt < 13727 Then
        x1 = 5 + ((((tt - 6528) Mod 25) Mod 5) * 5)
        x2 = 5 + ((((tt - 6528) Mod 25) \ 5) * 5)
        tt = tt - 6528
        tt = tt \ 25
    ElseIf tt > 13728 Then
        x2 = 5 + (((tt - 13728) Mod 5) * 5)
        x1 = 15
        tt = tt - 13728
        tt = (tt \ 5) + 288
    End If
            
    If tt < 192 Then
        b = ttbl(tt)
        t = b Mod 48
        d = b \ 48
    Else
        t = 47 - ((tt - 192) Mod 48)
        d = tt \ 48
    End If
    
    tp = t
    dp = d
    x1p = x1
    x2p = x2
End Sub
    

Public Sub bit_shuffle(code As Long, tval As Long, dval As Long, cval As Long)
    Dim tt As Long
    Dim cc As Long
    Dim x1 As Long
    Dim x2 As Long
    Dim nn As Long
    Dim top5 As Long
    Dim bot3 As Long
    Dim remainder As Long
    Dim t As Long
    Dim d As Long
    Dim outtime As Long
    Dim outdur As Long
    
    Dim sRemainder As String
    Dim sTop5 As String
    
    tt = 0
    cc = 0
    x1 = 0
    x2 = 0
    nn = code - 1
    top5 = nn \ 1000
    bot3 = nn Mod 1000
    remainder = bot3 Mod 32
    
    sTop5 = Convert(CStr(top5), 10, 2)
    sRem = Convert(CStr(remainder), 10, 2)
    
    tt = Bit(sTop5, 15) & Bit(sTop5, 14) & Bit(sTop5, 13) & Bit(sTop5, 11) & Bit(sTop5, 10) & Bit(sTop5, 7) & Bit(sTop5, 6) & Bit(sTop5, 5) & Bit(sTop5, 4) & Bit(sTop5, 3) & Bit(sTop5, 0) & Bit(sRem, 4) & Bit(sRem, 2) & Bit(sRem, 0)
    cc = Bit(sTop5, 16) & Bit(sTop5, 12) & Bit(sTop5, 8) & Bit(sTop5, 6) & Bit(sTop5, 2) & Bit(sRem, 1) & Bit(sRem, 3) & Bit(sRem, 1)
    
    If Bit(top5, 16) = "1" Then
        SetBit tt, 11, Bit(sTop5, 12)
        SetBit tt, 12, Bit(sTop5, 13)
        SetBit tt, 13, Bit(sTop5, 14)
        SetBit cc, 6, Bit(sTop5, 15)
    End If
    
    twiddle_tt tt, t, d, x1, x2
    
    outtime = (30 * t) + x2
    outdur = ((d + 1) * 30) - x1
    
    cval = cc + 1
    dval = outdur
    tval = outtime
    
End Sub

Public Sub map_top(year As Long, month_out As Long, day_out As Long, sTop5 As String, remainder As Long, mtoutp As Long, remoutp As Long)
    Dim year_mod16 As Long
    Dim year_mod100 As Long
    Dim ndigits As Long
    Dim nd As Long
    Dim flag7 As Long
    Dim j As Long
    Dim k As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim t3 As Long
    Dim ym As Long
    Dim datum As Long
    Dim mtout As Long
    Dim month_today As Long
    Dim year_today As Long
    Dim n As String
    Dim lCodeSize As Long
    
    month_today = month_out
    year_today = year
    
    year_mod100 = year Mod 100
    year_mod16 = year_mod100 Mod 16
    
    lCodeSize = Len(CStr(Val(sTop5)))
    
    nd = ndigits - 3 + 1
    
    remainder = (remainder + AddDigits(sTop5)) Mod 32
    
    If ndigits <= 6 Then
        For k = 0 To year_mod16
            Mid$(n, nd, 1) = CStr((Val(Mid$(n, nd, 1)) + day_out) Mod 10)
            n = ApplyKey(n, 2)
            remainder = remainder + Left$(n, 1)
        Next
        remainder = (remainder + (day_out * (month_today + 1))) Mod 32
    Else
        flag7 = Abs(ndigits = 7)
        ym = (year_today * 12) + month_today
        
        For k = 1 To 31
            t1 = (ym + 310 - k) Mod 31
            t1 = vOverSixTable(t1)
            t2 = (t1 Mod 16) - flag7
            t3 = ((t1 \ 16) Mod 16) - flag7
            
            If t2 = 0 And t3 < 3 Then
                t2 = 4
            ElseIf t3 = 0 Then
                t3 = IIf(t2 >= 3, 2, 4)
            End If
            t1 = Val(Mid$(n, t2 - 1, 1)) + (10 * Val(Mid$(n, t3 - 1, 1)))
            
            Do
                datum = vTtbl(t1) - ym
                While datum < 0 Or datum > 192
                    datum = datum + 192
                Wend
                If datum > 99 Then
                    t1 = datum
                End If
            Loop While datum > 99
            
            If t2 >= 0 And t2 < 5 And t3 >= 0 And t3 < 5 Then
                Mid$(n, t2, 1) = datum Mod 10
                Mid$(n, t3, 1) = datum \ 10
            Else
                MsgBox "Internal table index wild."
            End If
        Next
        mtout = Val(n)
        mtoutp = mtout
        remoutp = remainder
    End If
    
End Sub

Public Function ApplyKey(sNumber As String, lRepeat As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    ApplyKey = sNumber
    For lIndex = 1 To lRepeat
        Do
            For lPos = Len(ApplyKey) - 1 To 1 Step -1
                Mid$(ApplyKey, lPos, 1) = CStr((Val(Mid$(ApplyKey, lPos, 1)) + Val(Mid$(ApplyKey, lPos + 1, 1))) Mod 10)
            Next
        Loop Until Left$(ApplyKey, 1) <> "0"
    Next
End Function

Public Function ApplyInverseKey(sNumber As String, lRepeat As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    For lIndex = 1 To lRepeat
        Do
            ApplyInverseKey = sNumber
            For lPos = 1 To Len(scode) - 1
                Mid$(ApplyInverseKey, lPos, 1) = CStr((Val(Mid$(ApplyInverseKey, lPos, 1)) - Val(Mid$(ApplyInverseKey, lPos + 1, 1)) + 10) Mod 10)
            Next
        Loop Until Left$(ApplyInverseKey, 1) <> "0"
    Next
End Function

Public Sub decode_main(ByVal month_today As Long, ByVal day_today As Long, ByVal year_today As Long, ByVal sNewspaper As String, ByRef day_ret As Long, ByRef channel_ret As Long, ByRef starttime_ret As Long, ByRef duration_ret As Long)
    Dim sDescrambled As String
    Dim sBot3 As String
    Dim sTop5 As String
    Dim quo As Long
    Dim remainder As Long
    Dim mtout As Long
    Dim tval As Long
    Dim dval As Long
    Dim cval As Long
    Dim lDayOut As Long
    Dim channel_out As Long
    Dim starttime_out, duration_out As Long
    Dim modnews As Long
    Dim day_out As Long
    
    Initialise
    year_today = year_today Mod 100
    If (month_today < 1 Or month_today > 12) Then
        MsgBox "Invalid month"
        'usagex ();
    End If
    
    If (day_today < 1 Or day_today > 31) Then
        MsgBox "Invalid day of the month\n"
        'usagex ();
    End If
    
    If (sNewspaper < 1) Then
        MsgBox "DON'T TRY NUMBERS LESS THAN 1!\n"
        'usagex ();
    End If
    
    sDescrambled = Pad(ApplyKey(sNewspaper, 2), 8)
    
    sTop5 = Left$(sDescrambled, 5)
    sBot3 = Right$(sDescrambled, 3)
    
    quo = (Val(sBot3) - 1) \ 32
    remainder = (Val(sBot3) - 1) Mod 32
    lDayOut = quo + 1
    
    map_top year_today, month_today, lDayOut, sTop5, remainder, mtout, remainder

    modnews = mtout * 1000
    modnews = modnews + (day_out * 32) + remainder - 31

    bit_shuffle modnews, tval, dval, cval

    starttime_out = tval
    duration_out = dval
    channel_out = cval

    day_ret = day_out
    channel_ret = channel_out
    starttime_ret = starttime_out
    duration_ret = duration_out
End Sub

Public Sub Test()
    Dim lDay As Long
    Dim lChannel As Long
    Dim lStartTime As Long
    Dim lDuration As Long
    
    decode_main 3, 2, 2005, 73, lDay, lChannel, lStartTime, lDuration
    
End Sub

Public Sub main()
'
End Sub
