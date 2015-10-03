VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
   
Private Sub Form_Activate()
    DrawShape
    Scan
End Sub

Private Sub DrawShape()
    CurrentX = 110
    CurrentY = 10
    
    Print "O"
End Sub

Private Sub Scan()
    Dim x&
    Dim y&
    Dim offsetx&
    Dim offsety&
    Dim colour&
    Dim scolour&
    
    Randomize
    
    offsetx& = 100
    Printer.ScaleMode = vbPixels
    For x = 0 To 100
        For y = 0 To 100
            colour& = IIf(Int(Rnd * 2) = 0, 0, &HFFFFFF)
            scolour& = GetPixel(Me.hDC, x + offsetx&, y + offsety&)
            SetPixelV Me.hDC, x, y, colour&
            SetPixelV Me.hDC, x + offsetx&, y + offsety&, IIf(scolour <> 0&, colour&, &HFFFFFF - colour&)
            Printer.PSet (x, y), colour&
            Printer.PSet (x + offsetx&, y + offsety&), IIf(scolour <> 0&, colour&, &HFFFFFF - colour&)
        Next
    Next

    Printer.EndDoc
End Sub

