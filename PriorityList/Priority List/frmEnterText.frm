VERSION 5.00
Begin VB.Form frmEnterText 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Note"
   ClientHeight    =   285
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNote 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmEnterText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOriginalText As String

Private Sub Form_Activate()
    sOriginalText = txtNote.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim$(txtNote.Text) = "" Then
            txtNote.Text = sOriginalText
        End If
        Me.Hide
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 5 And UnloadMode <> 1 Then
        If Trim$(txtNote.Text) = "" Then
            txtNote.Text = sOriginalText
        End If
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    txtNote.Width = Me.ScaleWidth
    txtNote.Height = Me.ScaleHeight
End Sub


