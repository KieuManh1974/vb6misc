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

Private Sub Form_Load()
    Dim bByte As Byte
    Dim x As Long
    
    Open "D:\Media\Video\Captured\CHANNEL5 2004-01-28 2300-0105a.avi" For Random As #1
    Open "D:\Media\Video\Captured\test.avi" For Random As #2
    
    Seek #1, 1
    Seek #2, 1
    For x = 1 To 1000000
    'While Not EOF(2)
        Get #1, , bByte
        Put #2, , bByte
    'Wend
    Next
    Close #1
    Close #2
End Sub
