VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMemory 
      Height          =   8520
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   45
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim oHFEngine As New HFEngine
    Dim X As Long
    
    DoEvents
    DoEvents
    
    oHFEngine.Memory = RandomString
    Do
        oHFEngine.Execute
        txtMemory.Text = oHFEngine.Memory & " " & oHFEngine.PC & " " & oHFEngine.MC
        DoEvents
    Loop
End Sub

Private Function RandomString() As String
    Dim i As Long
    Randomize
    For i = 1 To 50
        'RandomString = RandomString & CStr(i Mod 2)
        RandomString = RandomString & CStr(Int(Rnd * 2))
    Next
End Function


