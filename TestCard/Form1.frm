VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
'    Line (0, 0)-(0, 767), vbWhite
'    Line (0, 0)-(1359, 0), vbWhite
    Line (1359, 0)-(1359, 767), vbWhite
    Line (0, 767)-(1359, 767), vbWhite
    
    Dim x As Single
    Dim y As Single
    
    For x = 0 To 1359 Step 16
        Line (x, 0)-(x, 767), vbWhite
    Next
    
    For y = 0 To 767 Step 16
        Line (0, y)-(1359, y), vbWhite
    Next
End Sub

Private Sub Form_Click()
    Unload Me
End Sub
