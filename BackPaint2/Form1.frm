VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pic As StdPicture

Private Sub Form_Activate()
    Dim x As Long
    For x = 1 To 10000
        PaintPicture pic, Rnd * Me.Width, Rnd * Me.Height, (Int(Rnd * 2) * 2 - 1) * (566.9 * pic.Width / 1000) / (1 + Rnd * 5), (Int(Rnd * 2) * 2 - 1) * (566.9 * pic.Height / 1000) / (1 + Rnd * 5)
    Next
End Sub

Private Sub Form_Load()
'    Set pic = LoadPicture("c:\crop circles\16b.jpg")
'    Set pic = LoadPicture("c:\grass.bmp")
    Set pic = LoadPicture("c:\sk.bmp")
End Sub

 Private Sub BackPaint(x As Single, Y As Single, w As Single, h As Single)
     Dim xpos As Integer
     Dim xoffset As Integer
     Dim ypos As Integer
     Dim yoffset As Integer
     Dim pwidth As Single
     Dim pheight As Single

     Dim tppx As Integer
     Dim tppy As Integer

     tppx = Screen.TwipsPerPixelX
     tppy = Screen.TwipsPerPixelY

     x = x / tppx
     Y = Y / tppy
     w = w / tppx
     h = h / tppy

     pwidth = Int(0.5 + (566.9 * pic.Width / 1000) / Screen.TwipsPerPixelX)
     pheight = Int(0.5 + (566.9 * pic.Height / 1000) / Screen.TwipsPerPixelY)

     ypos = Y
     yoffset = ypos - (ypos \ pheight) * pheight

     While ypos + (pheight - yoffset) <= (Y + h - 1)
         xpos = x
         xoffset = xpos - (xpos \ pwidth) * pwidth
         While xpos + (pwidth - xoffset) <= (x + w - 1)
             PaintPicture pic, xpos * tppx, ypos * tppy, , , xoffset * tppx, yoffset * tppy
             xpos = xpos + (pwidth - xoffset)
             xoffset = 0
         Wend
         If (w + x - xpos) > 0 Then
             PaintPicture pic, xpos * tppx, ypos * tppy, , , xoffset * tppx, yoffset * tppy, (w + x - xpos) * tppx
         End If

         ypos = ypos + (pheight - yoffset)
         yoffset = 0
     Wend
     If (h + Y - ypos) > 0 Then
         xpos = x
         xoffset = xpos - (xpos \ pwidth) * pwidth
         While xpos + (pwidth - xoffset) <= (x + w - 1)
             PaintPicture pic, xpos * tppx, ypos * tppy, , , xoffset * tppx, yoffset * tppy, , (h + Y - ypos) * tppy
             xpos = xpos + (pwidth - xoffset)
             xoffset = 0
         Wend
         If (w + x - xpos) > 0 Then
             PaintPicture pic, xpos * tppx, ypos * tppy, , , xoffset * tppx, yoffset * tppy, (w + x - xpos) * tppx, (h + Y - ypos) * tppy
         End If
     End If

 End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        BackPaint x, Y, 1000, 1000
    End If
    
End Sub
