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

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private mlTimerID As Long

Sub main()
    mlTimerID = SetTimer(0, 0, 1000, AddressOf ResetTime)
    'ResetTime
End Sub

Sub ResetTime()
    Dim stTime As SYSTEMTIME
    Dim dTime As Double
    Dim dDate As Date

    GetSystemTime stTime
    With stTime
        dTime = (CDbl(DateSerial(.wYear, .wMonth, .wDay)) * 86400 + CDbl(.wHour) * 3600 + CDbl(.wMinute) * 60 + CDbl(.wSecond)) * 1000 + CDbl(.wMilliseconds) + 2000
        dDate = (CDate(dTime / 1000 / 86400))
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
        .wMilliseconds = dTime - Int(dTime / 1000) * 1000
    End With
    SetSystemTime stTime
End Sub

Public Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Public Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCount
    GetCounter = Format$((cCount - cCounters(lCounterIndex) - CCur(0.0007)) / cFrequency, "0.000000000")
End Function

