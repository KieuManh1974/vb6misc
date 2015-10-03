Attribute VB_Name = "Clock"
Option Explicit

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private cFrequency As Currency
Private cCounters(0 To 5) As Currency

Public stSavedTime As SYSTEMTIME

Public Sub SetTime(ByVal dSetTime As Date)
    Dim stTime As SYSTEMTIME
    Dim stTime2 As SYSTEMTIME
    Dim dTime As Double
    Dim dSec As Double
    Dim dMilli As Double
    
    StartCounter
    GetSystemTime stSavedTime
    
    With stTime
        .wYear = Year(dSetTime)
        .wMonth = Month(dSetTime)
        .wDay = Day(dSetTime)
        .wHour = Hour(dSetTime)
        .wMinute = Minute(dSetTime)
        .wSecond = Second(dSetTime)
        .wMilliseconds = 0
    End With
    SetSystemTime stTime

End Sub

Public Sub RestoreTime()
    Dim dTime As Double
    Dim dSec As Double
    Dim dMilli As Double
    Dim dDay As Double
    Dim tDate As Date
    Dim iAdjust As Integer
    
    dTime = GetCounter()
    dSec = Int(dTime)
    dMilli = (dTime - dSec) * 1000
    
    With stSavedTime
        .wMilliseconds = .wMilliseconds + dMilli
        .wSecond = .wSecond + Adjust(.wMilliseconds, 1000) + dSec
        .wMinute = .wMinute + Adjust(.wSecond, 60)
        .wHour = .wHour + Adjust(.wMinute, 60)

        tDate = CDbl(DateSerial(.wYear, .wMonth, .wDay)) + Adjust(.wHour, 24)
        .wDay = Day(tDate)
        .wMonth = Month(tDate)
        .wYear = Year(tDate)
    End With
    SetSystemTime stSavedTime
End Sub

Public Function Adjust(lValue As Integer, ByVal lMod As Integer) As Integer
    While lValue >= lMod
        lValue = lValue - lMod
        Adjust = Adjust + 1
    Wend
End Function

Private Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Private Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCount
    GetCounter = Format$((cCount - cCounters(lCounterIndex) - CCur(0.0007)) / cFrequency, "0.000000000")
End Function

