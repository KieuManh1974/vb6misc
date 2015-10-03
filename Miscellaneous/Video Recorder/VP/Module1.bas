Attribute VB_Name = "Module1"
Option Explicit


Public Sub split_digits(n As Integer, a As String)

End Sub

'void split_digits (int n, unsigned char *a) {
'    int             i;
'    unsigned char   digit;
'
'    clear_ndigits (a);
'
'    for (i = 0; i < NDIGITS; i++) {
'        digit = n % 10;
'        a[i] = digit;
'        n = (n - digit) / 10;
'    }
'}
'
'int
'count_digits (int val)
'{
'    int             ndigits;
'    if (val < 0) {
'        printf ("Error: code 0 or negative\n");
'        ndigits = 0;
'    } else if (val < 1)
'        ndigits = 0;
'    else if (val < 10)
'        ndigits = 1;
'    else if (val < 100)
'        ndigits = 2;
'    else if (val < 1000)
'        ndigits = 3;
'    else if (val < 10000)
'        ndigits = 4;
'    else if (val < 100000)
'        ndigits = 5;
'    else if (val < 1000000)
'        ndigits = 6;
'    else if (val < 10000000)
'        ndigits = 7;
'    else if (val < 100000000)
'        ndigits = 8;
'    Else
'        ndigits = 9;
'
'    if (ndigits > 8) {
'        printf ("ERROR: %d has more than 8 digits (it has %d digits)\n", val, ndigits);
'        usagex ();
'    }
'    return ndigits;
'}

Public Function func1(code As Long) As Integer
    Dim x As Integer
    Dim nd As Integer
    Dim a(ndigits) As String * 1
    Dim sum As Integer
    Dim i, j As Integer
    Dim ndigits As Integer

    x = code
    
    split_digits x, a
    ndigits = count_digits(x)
    nd = ndigits - 1
    
    Do
        i = 0
        Do
            j = 1
            If nd >= 1 Then
                Do
                    a(j) = (a(j - 1) + a(j)) Mod 10
                    j = j + 1
                While j <= nd
            End If
            i = i + 1
        While i <= 2
    While a(nd) = 0
    
    sum = 0
    j = 1
    For i = 0 To ndigits
        sum = sum + j * a(i)
        j = j * 10
    Next
End Function


Public Sub decode_main(ByVal month_today As Integer, ByVal day_today As Integer, ByVal year_today As Integer, ByVal newspaper As Integer, ByRef day_ret As Integer, ByRef channel_ret As Integer, ByRef starttime_ret As Integer, ByRef duration_ret As Integer)
    Dim s1_out, bot3, top5, quo, remainder As Integer
    Dim mtout, tval, dval, cval As Integer
    Dim day_out, channel_out As Integer
    Dim starttime_out, duration_out As Integer
    Dim modnews As Integer

    year_today = year_today Mod 100
    If (month_today < 1 Or month_today > 12) Then
        MsgBox "Invalid month"
        'usagex ();
    End If
    
    If (day_today < 1 Or day_today > 31) Then
        MsgBox "Invalid day of the month\n"
        'usagex ();
    End If
    
    If (newspaper < 1) Then
        MsgBox "DON'T TRY NUMBERS LESS THAN 1!\n"
        'usagex ();
    End If
    If (g_iflag) Then
        s1_out = newspaper
    Else
        s1_out = func1(newspaper)
    End If
    
    top5 = s1_out / 1000
    bot3 = (s1_out Mod 1000)
    quo = (bot3 - 1) / 32
    remainder = (bot3 - 1) Mod 32
    day_out = quo + 1
    
    map_top year_today, month_today, day_out, top5, remainder, mtout, remainder

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

