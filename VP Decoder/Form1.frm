VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDecode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2430
      Width           =   7755
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   240
      MaxLength       =   8
      TabIndex        =   0
      Top             =   315
      Width           =   7635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ApplyKey(sNumber As String, lRepeat As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    If Left$(sNumber, 1) = "0" Then
        ApplyKey = ""
        Exit Function
    End If
    
    ApplyKey = sNumber
    For lIndex = 1 To lRepeat
        Do
            For lPos = Len(ApplyKey) - 1 To 1 Step -1
                Mid$(ApplyKey, lPos, 1) = CStr((Val(Mid$(ApplyKey, lPos, 1)) + Val(Mid$(ApplyKey, lPos + 1, 1))) Mod 10)
            Next
        Loop Until Left$(ApplyKey, 1) <> "0"
    Next
End Function

Public Function ApplyInverseKey(sNumber As String, lRepeat As Long) As String
    Dim lPos As Long
    Dim lIndex As Long
    
    If Left$(sNumber, 1) = "0" Then
        ApplyInverseKey = ""
        Exit Function
    End If
            
    ApplyInverseKey = sNumber
    For lIndex = 1 To lRepeat
        Do
    
            For lPos = 1 To Len(ApplyInverseKey) - 1
                Mid$(ApplyInverseKey, lPos, 1) = CStr((Val(Mid$(ApplyInverseKey, lPos, 1)) - Val(Mid$(ApplyInverseKey, lPos + 1, 1)) + 10) Mod 10)
            Next
        Loop Until Left$(ApplyInverseKey, 1) <> "0"
    Next
End Function

Private Sub txtCode_Change()
    txtDecode.Text = ApplyKey(txtCode.Text, 3)
End Sub

Private Sub txtDecode_Change()
    txtCode.Text = ApplyInverseKey(txtDecode.Text, 3)
End Sub
