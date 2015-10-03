VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkZero 
      Caption         =   "With Zero"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkSpacing 
      Caption         =   "Spacing 2"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkSpacing 
      Caption         =   "Spacing 1"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkSpacing 
      Caption         =   "Spacing 0"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton optParity 
      Caption         =   "Both Odd"
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton optParity 
      Caption         =   "Odd && Even"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton optParity 
      Caption         =   "Both Even"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CheckBox chkCarry 
      Caption         =   "Allow Carry"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check2_Click(Index As Integer)

End Sub

Private Sub Form_Load()

End Sub
