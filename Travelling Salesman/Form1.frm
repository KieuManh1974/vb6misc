VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   897
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type Point
    X As Long
    Y As Long
End Type

Private Points() As Point

Private Sub Form_Activate()
    Dim X As Long
    Dim Y As Long
    Dim d1 As Double
    Dim d2 As Double
    
    ReDim Points(3) As Point
    
    Points(0) = NewPoint(100, 270)
    Points(1) = NewPoint(300, 200)
    Points(2) = NewPoint(300, 300)
    
    For Y = 0 To 500
        For X = 0 To 1000
            Points(3).X = X
            Points(3).Y = Y
            d1 = Distance(Points(0), Points(1)) + Distance(Points(1), Points(2)) + Distance(Points(2), Points(3))
            d2 = Distance(Points(0), Points(2)) + Distance(Points(2), Points(1)) + Distance(Points(1), Points(3))
            If d1 < d2 Then
                SetPixelV Me.hdc, X, Y, vbRed
            Else
                SetPixelV Me.hdc, X, Y, vbBlue
            End If
        Next
    Next
    
    For X = 0 To 3
        SetPixelV Me.hdc, Points(X).X, Points(X).Y, vbWhite
    Next
End Sub

Private Function NewPoint(ByVal X As Long, ByVal Y As Long) As Point
    NewPoint.X = X
    NewPoint.Y = Y
End Function

Private Function Distance(p1 As Point, p2 As Point) As Double
    Distance = Sqr((p2.X - p1.X) * (p2.X - p1.X) + (p2.Y - p1.Y) * (p2.Y - p1.Y))
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Caption = X & "," & Y
End Sub
