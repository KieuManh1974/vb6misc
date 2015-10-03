VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00800000&
   ForeColor       =   &H0000FF00&
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

Private Sub Form_Activate()
    Dim x As Long
    Dim y As Long
    
    
    Me.FillStyle = vbSolid
    Me.DrawWidth = 20
    For x = -120 To 1400 Step 60
        For y = -200 To 1000 Step 70
            'Me.Circle (x, y + x / 10), 0
            Me.Line (x + y / 10, y + x / 10)-Step(20, -35)
        Next
    Next
End Sub

