VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8445
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
    Dim lNumber As Long
    Dim lMultiplier As Long
    Dim lColour As Long
    
    For lMultiplier = 3 To 31 Step 2
        lColour = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        For lNumber = 1 To 10000 Step 2
            PSet (lNumber * 2 + 100, Me.Height - 800 - Reduce(lNumber, lMultiplier) * 100), lColour
        Next
        
    Next
End Sub

