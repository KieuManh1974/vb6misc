VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Destination 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   3120
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Form_Activate()
    Dim blockX As Long
    Dim blockY As Long
    Dim centreX As Long
    Dim centreY As Long
    
    ScanForBlock 0, 1, blockX, blockY
    ScanBlock 0, 1, blockX, blockY, centreX, centreY
End Sub

Private Sub ScanForBlock(ByVal xdir As Long, ByVal ydir As Long, blockX As Long, blockY As Long)
    Dim picwidth As Long
    Dim picheight As Long
    Dim pix As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    
    picwidth = Int(0.5 + (566.9 * Source.Picture.Width / 1000) / Screen.TwipsPerPixelX)
    picheight = Int(0.5 + (566.9 * Source.Picture.Height / 1000) / Screen.TwipsPerPixelY)
    
    For y = 1 To picheight
       For x = 1 To picwidth
           pix = GetPixel(ByVal Source.hdc, ByVal x, ByVal y)
           i = pix Mod 256
           'SetPixelV ByVal Destination.hdc, ByVal x, ByVal y, ByVal -(i < 128) * vbWhite
           If i < 128 Then
               Exit For
           End If
       Next
       If i < 128 Then
           Exit For
       End If
    Next
    blockX = x
    blockY = y
End Sub

Private Sub ScanBlock(ByVal xdir As Long, ByVal ydir As Long, ByVal blockX As Long, ByVal blockY As Long, centreX As Long, centreY As Long)
    Dim x As Long
    Dim y As Long
    Dim xx As Long
    Dim pix As Long
    Dim area As Long
    Dim i  As Long
    Dim j As Long
    
    Dim sumcol() As Long
    Dim sumrow() As Long
    
    Dim leftmost As Long
    Dim rightmost As Long
    Dim upmost As Long
    Dim downmost As Long
    
    leftmost = 1000000#
    downmost = 1000000#
    upmost = blockY
    
    y = blockY
    x = blockX
    j = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
    While j < 128
        i = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
        While i < 128
            area = area + 1
            If x < leftmost Then
                leftmost = x
            End If
            x = x - 1
            i = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
        Wend
        
        x = blockX + 1
         i = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
        While i < 128
            area = area + 1
            If x > rightmost Then
                rightmost = x
            End If
            x = x + 1
            i = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
        Wend
        y = y + 1
        downmost = y
        x = blockX
        j = GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256
    Wend
    
    ReDim sumcol(leftmost - 2 To rightmost + 2) As Long
    ReDim sumrow(upmost - 2 To downmost + 2) As Long
    
    For x = leftmost - 2 To rightmost + 2
        For y = upmost - 2 To downmost + 2
            i = 255 - (GetPixel(ByVal Source.hdc, ByVal x, ByVal y) Mod 256)
            SetPixelV ByVal Destination.hdc, ByVal x, ByVal y, i
            sumcol(x) = sumcol(x) + i
            sumrow(y) = sumrow(y) + i
        Next
    Next
End Sub

