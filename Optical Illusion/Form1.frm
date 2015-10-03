VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
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
    Dim lIndex As Long
    
    For lIndex = 0 To 1
        SymbolSquare lIndex * 7, 0, 7, 7, lIndex * 2 + 1, Choose(lIndex + 1, vbRed, vbGreen)
    Next
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    
    Cls
    For lIndex1 = 0 To 1
        For lIndex2 = 0 To 1
            SymbolSquare lIndex1 * 7, lIndex2 * 7, 7, 7, ((lIndex1 + lIndex2) Mod 2) * 2 + 1, vbWhite
        Next
    Next
End Sub


Private Function SymbolSquare(ByVal lPosX As Long, ByVal lPosY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal sChar As String, ByVal lColour As Long)
    Dim lX As Long
    Dim lY As Long

    'BackColor = vbBlack
    ForeColor = lColour
    For lX = 0 To lWidth - 1
        For lY = 0 To lHeight - 1
            CurrentX = (lPosX + lX) * 10 + 30
            CurrentY = (lPosY + lY) * 10 + 30
            Print sChar
        Next
    Next
    
End Function
