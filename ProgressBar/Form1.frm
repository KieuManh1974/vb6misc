VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Min As Long
Public Max As Long
Public Value As Long
Public BarWidth As Single
Public BarHeight As Single
Public Colour As Long
Public BackColour As Long

Private Type RGBSplit
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
End Type


Private Sub Form_Activate()
    Dim y As Long
    
    Min = 0
    Max = 1000
    Value = 57
    BarWidth = 100
    BarHeight = 20
    Colour = vbRed
    BackColour = &HE0E0E0
    
    For Value = 0 To 1000
        DrawBar
        For y = 0 To 1000
            DoEvents
        Next
    Next
End Sub

Private Sub DrawBar()
    Dim sWidth As Single
    Dim sRemainder As Single
    Dim lSubColour As RGBSplit
    Dim lSubBackColour As RGBSplit
    
    sWidth = Int(BarWidth * (Value - Min) / (Max - Min))
    sRemainder = (BarWidth * (Value - Min) / (Max - Min)) - sWidth
    Line (100, 100)-Step(sWidth, BarHeight), Colour, BF
    CopyMemory lSubColour, Colour, 4&
    CopyMemory lSubBackColour, BackColour, 4&
    Line (100 + sWidth, 100)-Step(0, BarHeight), RGB((CLng(lSubColour.Red) - lSubBackColour.Red) * sRemainder + lSubBackColour.Red, (CLng(lSubColour.Green) - lSubBackColour.Green) * sRemainder + lSubBackColour.Green, (CLng(lSubColour.Blue) - lSubBackColour.Blue) * sRemainder + lSubBackColour.Blue), BF
End Sub
