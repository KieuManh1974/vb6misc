Attribute VB_Name = "Performance"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private cFrequency As Currency
Private cCounters(0 To 5) As Currency


Private N As Long
Private Z As Long
Private A As Long
Private V As Long
Private C As Long

Public Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Public Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCount
    GetCounter = Format$((cCount - cCounters(lCounterIndex) - CCur(0.0008)) / cFrequency, "0.000000000")
End Function

Sub Main2()
    Dim lCount As Long
    Dim lCount2 As Long
    Dim dTime As Double
    Dim dFastestTime As Double
    Const lIterations As Long = 100000
    
                                    Dim lLoNibble As Long
                                    Dim lHiNibble As Long
                                    Dim lSum As Long
                                    Dim lAdjusted As Long
                                    Dim lALoNibble As Long
                                    Dim lHalfCarry As Long
                                    Dim lHiNibbleOverTen As Long
                                    
                                    Dim lLocation As Long

    dFastestTime = 1000000#
    For lCount2 = 1 To 10
        StartCounter
        For lCount = 1 To lIterations

                                    lALoNibble = A And &HF
                                    lSum = A + gyMem(lLocation) + (lALoNibble >= 10) * (lALoNibble - 10) - 248&
                                    V = ((lSum >= -128) And (lSum <= 127)) * -64
                                    V = V And (((A Xor gyMem(lLocation)) And &H80) = 0)
                                    lLoNibble = lALoNibble + (gyMem(lLocation) And &HF) + C
                                    lHalfCarry = lLoNibble >= 10
                                    lHiNibble = ((A And &HF0) + (gyMem(lLocation) And &HF0)) - lHalfCarry * 16
                                    C = Abs(lHiNibble >= 160)
                                    A = (((lLoNibble + lHalfCarry * 10) And &HF)) + ((lHiNibble - C * 160) And &HF0)
                                    N = A And 128
                                    Z = (A = 0) * -2

        Next
        dTime = CDbl(GetCounter) / CDbl(lIterations)
        If dTime < dFastestTime Then
            dFastestTime = dTime
        End If
    Next

    Debug.Print Scientific(dFastestTime)
    If Dir(App.Path & "\perf.txt") <> "" Then Kill App.Path & "\perf.txt"
    Open App.Path & "\perf.txt" For Binary As #2
    Put #2, , Scientific(dFastestTime)
    Close #2
End Sub


Public Function Scientific(ByVal dValue As Double) As String
    Dim lMultiplier As Long
    Dim vNames As Variant
    
    lMultiplier = 5
    vNames = Array("peta", "tera", "giga", "mega", "kilo", "", "milli", "micro", "nano", "pico", "femto")
    If Abs(dValue) < 1 Then
        While Abs(dValue) < 1
            dValue = dValue * 1000
            lMultiplier = lMultiplier + 1
        Wend
    ElseIf Abs(dValue) >= 1000 Then
        While Abs(dValue) >= 1000
            dValue = dValue / 1000
            lMultiplier = lMultiplier - 1
        Wend
    End If
    
    Scientific = Format$(dValue, "0.000") & " " & vNames(lMultiplier)
End Function
