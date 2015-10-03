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

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private cFrequency As Currency
Private cCounters(0 To 5) As Currency

Private Sub Form_Load()
    Dim yMemory(-7 To 100) As Byte
    Dim lAddress As Long
    Dim lRegister As Long
    
    lRegister = -1
    SetFrequency
    StartCount
    lAddress = yMemory(-7) + yMemory(-6) * 256 + yMemory(lRegister)
    lAddress = yMemory(lAddress) + yMemory(lAddress + 1) * 256
    Debug.Print GetCount
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

