VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Binary(ByVal fValue As Double, ByVal lPrecision As Long) As String
    fValue = fValue - Int(fValue)
    
    While lPrecision > 0
        fValue = fValue * 2
        Binary = Binary & Int(fValue)
        fValue = fValue - Int(fValue)
        lPrecision = lPrecision - 1
    Wend
End Function

Private Sub Timer1_Timer()
    Caption = Binary(Time(), 16)
End Sub

Private Function SplitBinary(ByVal x As Long)

End Function
