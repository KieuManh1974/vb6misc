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
   Begin VB.TextBox txtOutput 
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
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   6135
   End
   Begin VB.TextBox txtInput 
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

Private sTxtOutputHeight As Single

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim oParseTree As ParseTree
    Dim oSetDef As SetDef
    
    If KeyAscii = 10 Then
        KeyAscii = 0
        Stream.Text = "{" & Replace$(txtInput.Text, vbCrLf, "") & "}"
        Set oParseTree = New ParseTree
        If oParser.Parse(oParseTree) Then
            Set AllSets = New Dictionary
            Set oSetDef = New SetDef
            oSetDef.CreateSet oParseTree
            txtOutput.Text = oSetDef.AllText
        Else
            txtOutput.Text = "error"
        End If
    End If
End Sub

Private Sub Form_Load()
    sTxtOutputHeight = Me.ScaleHeight - txtOutput.Height
    InitialiseParser
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtInput.Width = Me.ScaleWidth
    txtOutput.Width = Me.ScaleWidth
    txtOutput.Height = Me.ScaleHeight - sTxtOutputHeight
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then
        txtInput.Text = Compress(txtOutput.Text)
        KeyAscii = 0
    End If
End Sub

