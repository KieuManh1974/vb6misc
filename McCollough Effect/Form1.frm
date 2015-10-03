VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim lX As Long
    Dim lY As Long
    Dim lXDiff As Long
    
    lXDiff = 2
    While lX < 700
        Line (lX, 0)-Step(1, 1040), vbGreen, BF
        lX = lX + lXDiff
        lXDiff = lXDiff + 1
        If lXDiff > 10 Then
            lXDiff = 2
        End If
    Wend
    
    
    For lY = 0 To 1040 Step 8
        Line (700, lY)-Step(700, 3), vbMagenta, BF
    Next
End Sub

Private Sub Form_Click()
    Unload Me
End Sub


