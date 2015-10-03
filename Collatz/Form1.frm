VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
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
    Dim X As Long
    Dim Y As Long
    
    For X = 3 To 110000# Step 2
        Me.PSet (Log(X) / Log(2) * 200 - 2400, Me.ScaleHeight - CollatzSteps(X) * 2), vbBlack
        'Me.PSet (Log(x) / Log(2) * 10, Me.ScaleHeight - ZeroCount(x) * 20), vbBlack
        'Me.PSet (ZeroCount(x) * 20, Me.ScaleHeight - CollatzSteps(x)), vbBlack
        'Me.PSet (X, Me.ScaleHeight - CollatzSteps(X) * 2), vbBlack
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Line (0, 0)-Step(100, 50), vbWhite, BF
    CurrentX = 0
    CurrentY = 0
    Print Log(X) / Log(2)
    Print (ScaleHeight - Y) * 2
End Sub
