VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FSCompare"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSnapshot2 
      Caption         =   "Snapshot 2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSnapshot1 
      Caption         =   "Snapshot 1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oScan As New Snapshot

Private Sub cmdSnapshot1_Click()
    cmdSnapshot1.Enabled = False
    oScan.Scan 1
    cmdSnapshot1.Enabled = True
End Sub

Private Sub cmdSnapshot2_Click()
    cmdSnapshot2.Enabled = False
    oScan.Scan 2
    cmdSnapshot2.Enabled = True
End Sub

Private Sub cmdCompare_Click()
    cmdCompare.Enabled = False
    oScan.Compare
    cmdCompare.Enabled = True
End Sub

