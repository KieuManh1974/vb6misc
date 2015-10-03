VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTime 
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   2745
      Width           =   2295
   End
   Begin VB.TextBox txtPeriod 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Text            =   "1"
      Top             =   1935
      Width           =   2295
   End
   Begin VB.TextBox txtSemiminorAxis 
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Text            =   "2"
      Top             =   1170
      Width           =   2295
   End
   Begin VB.TextBox txtSemimajorAxis 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "2"
      Top             =   405
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   2475
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Period"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1665
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Semi Minor Axis"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Semi Major Axis"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

' Find the angle given the major and minor semi-axes
Private Function AngleFromTime(ByVal fTime As Double, ByVal fMajorSemiAxis As Double, ByVal fMinorSemiAxis As Double) As Double
    
End Function
