VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
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

Private mlBlocks() As Long
Private Const mlMax As Long = 300
Private Const mlBlockSize As Long = 26

Private Sub Form_Load()
    ReDim mlBlocks(mlMax)
    SetBlocks 2
    SetBlocks 3
    'SetBlocks 5
    'SetBlocks 7
    DrawBlocks
End Sub

Private Sub SetBlocks(ByVal lStepSize As Long, Optional ByVal lOffset As Long)
    Dim lIndex As Long
    
    
    For lIndex = 0 To mlMax
        If mlBlocks(lIndex) = 2 Then
            mlBlocks(lIndex) = 1
        End If
    Next
    
    For lIndex = lOffset To mlMax Step lStepSize
        Select Case mlBlocks(lIndex)
            Case 0
                mlBlocks(lIndex) = 2
            Case 1
        End Select
    Next
End Sub

Private Sub DrawBlocks()
    Dim lIndex As Long
    Dim vColours As Variant
    Dim lX As Long
    Dim lY As Long
    
    vColours = Array(vbBlack, vbWhite, vbRed, vbBlue, vbCyan)
    
    For lIndex = 0 To mlMax
        lX = (lIndex Mod 48)
        lY = (lIndex \ 48)
        
        Me.Line (lX * mlBlockSize, 50 + lY * mlBlockSize * 3)-Step(mlBlockSize - 1, mlBlockSize - 1), vColours(mlBlocks(lIndex)), BF
        Me.PSet (lX * mlBlockSize, 50 + mlBlockSize + lY * mlBlockSize * 3), vbBlack
        Print lIndex
    Next
End Sub
