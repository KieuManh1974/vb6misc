Attribute VB_Name = "Module1"
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

Sub main()
    Dim stTime As SYSTEMTIME
    Dim stTime2 As SYSTEMTIME
    Dim dTime As Double
    Dim dSec As Double
    Dim dMilli As Double
    
    StartCounter
    GetSystemTime stTime2
    
    With stTime
        .wYear = 2005
        .wMonth = 2
        .wDay = 17
        .wHour = 11
        .wMinute = 41
        .wSecond = 0
        .wMilliseconds = 0
    End With
    SetSystemTime stTime
    
    dTime = GetCounter()
    dSec = Int(dTime)
    dMilli = (dTime - dSec) * 1000
    
    With stTime2
        .wMilliseconds = .wMilliseconds + dMilli
        .wSecond = .wSecond + dSec + Int(.wMilliseconds / 1000)
        .wMinute = .wMinute + Int(.wSecond / 60)
        .wHour = .wHour + Int(.wMinute / 60)
        .wDay = .wDay + Int(.wHour / 24)
        .wMonth = .wMonth
        .wYear = .wYear + Int(.wMonth / 12)
    End With
End Sub

Public Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Public Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCount
    GetCounter = Format$((cCount - cCounters(lCounterIndex) - CCur(0.0007)) / cFrequency, "0.000000000")
End Function


