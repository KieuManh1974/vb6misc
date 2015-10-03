VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SplitFiles"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGB 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtKB 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtBytes 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtMB 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "&Split"
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "GB"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "KB"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Bytes"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "MB"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fragment Size"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lFragmentSize As Long

Private Sub cmdJoin_Click()
    JoinFilesSub
    MsgBox "Files have been joined."
End Sub

Private Sub cmdSplit_Click()
    If lFragmentSize <> 0 Then
        SplitFilesSub lFragmentSize
        MsgBox "Files have been split."
    Else
        MsgBox "Please choose a fragment size"
    End If
End Sub

Private Sub txtBytes_Change()
    lFragmentSize = Val(txtBytes.Text)
End Sub

Private Sub txtKB_Change()
    lFragmentSize = 2 ^ 10 * Val(txtKB.Text)
End Sub

Private Sub txtMB_Change()
    lFragmentSize = 2 ^ 20 * Val(txtMB.Text)
End Sub

Private Sub txtGB_Change()
    lFragmentSize = 2& ^ 30& * Val(txtGB.Text)
End Sub
