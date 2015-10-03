VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   615
      Left            =   3420
      TabIndex        =   2
      Top             =   2280
      Width           =   1155
   End
   Begin VB.TextBox txtOutput 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   4455
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRead_Click()
    Dim oNC As New NumberCruncher.EnglishNumber
    Dim dOutput As String
    
    If oNC.EnglishToNumber(txtInput.Text, dOutput) Then
        txtOutput.Text = dOutput
    Else
        txtOutput.Text = "This is not a number"
    End If
End Sub
