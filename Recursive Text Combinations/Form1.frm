VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Combinations"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   6135
   End
   Begin VB.TextBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
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
    Dim oListSet As ListSet
    
    If KeyAscii = 10 Then
        KeyAscii = 0
        ParserTextString.ParserText = Combo1.Text
        Set oParseTree = New ParseTree
        If oParser.Parse(oParseTree) Then
            Set oListSet = CreateStructure2(oParseTree)
            Text1.Text = oListSet.Text
        End If
    End If
End Sub

Private Sub Form_Load()
    Text1Height = Me.ScaleHeight - Text1.Height
    InitialiseParser
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Combo1.Width = Me.ScaleWidth
    Text1.Width = Me.ScaleWidth
    Text1.Height = Me.ScaleHeight - Text1Height
End Sub

