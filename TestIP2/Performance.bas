Attribute VB_Name = "Performance"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private cFrequency As Currency
Private cCounters(0 To 5) As Currency

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
