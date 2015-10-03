VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   6480
      Width           =   3840
   End
   Begin VB.TextBox txtColour 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4920
      Width           =   2655
   End
   Begin VB.PictureBox pctColour 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" _
  (ByVal hDC As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal crColor As Long) As Byte
   
Private Declare Function GetPixel Lib "gdi32" _
  (ByVal hDC As Long, _
   ByVal X As Long, _
   ByVal Y As Long) As Long
   
Private Sub Form_Load()
    PaintTriangle 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lColour As Long
    
    lColour = GetPixel(Me.hDC, X, Y)
    If lColour <> -1 Then
        pctColour.BackColor = lColour
        txtColour.Text = "#" & Hx(lColour Mod 256) & Hx(lColour \ 256 Mod 256) & Hx(lColour \ 65536 Mod 256)
    End If
End Sub

Private Function Hx(lNumber As Long) As String
    Hx = Right$(Hex$(lNumber + 256), 2)
End Function

Private Sub HScroll1_Change()
    PaintTriangle HScroll1.Value
End Sub

Private Sub PaintTriangle(ByVal lIntensity As Long)
    Dim lX As Long
    Dim lY As Long
    Dim lLower As Long
    Dim lUpper As Long
    
    lIntensity = lIntensity / 3
    
    lLower = lIntensity
    lUpper = 255 - lIntensity
    
    Cls
    For lY = lLower To lUpper
        For lX = lLower To lUpper - lY
            SetPixelV Me.hDC, lX + lY \ 2, lY, RGB(2 * lIntensity + lUpper - lY - lX, 2 * lIntensity + lX, 2 * lIntensity + lY)
        Next
    Next
End Sub
