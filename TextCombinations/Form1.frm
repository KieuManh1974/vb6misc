VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Combinations"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Combo1 
      Height          =   1965
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Text1Height As Single

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim oParseTree As ParseTree
    
    If KeyAscii = 10 Then
        KeyAscii = 0
        ParserTextString.ParserText = Combo1.Text
        Set oParseTree = New ParseTree
        If oParser.Parse(oParseTree) Then
            Text1.Text = CreateStructure(oParseTree)
        End If
    End If
End Sub

Private Sub Form_Load()
    Text1Height = Me.ScaleHeight - Text1.Height
    InitialiseParser
End Sub

Private Sub Form_Resize()
    Combo1.Width = Me.ScaleWidth
    Text1.Width = Me.ScaleWidth
    Text1.Height = Me.ScaleHeight - Text1Height
End Sub
