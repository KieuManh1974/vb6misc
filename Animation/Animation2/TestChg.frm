VERSION 5.00
Object = "*\APicChg.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   2895
   ClientTop       =   990
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   5940
   Begin VB.Frame Frame1 
      Caption         =   "Change Style"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5535
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Grow New after Shrinking Old"
         Height          =   495
         Index           =   9
         Left            =   3600
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Grid and Fill"
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Kick Bytes"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Window Shade"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Compress Old - Expand New"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   2895
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Slide Old Left Displaying New"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2895
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Slide Old Left Then New Right"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Shrink Old - Grow New"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Bring in Slats of new Image"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optChangeStyle 
         Caption         =   "Slide Left-to-Right--Rod's Example"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.TextBox txtSteps 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2475
      Index           =   1
      Left            =   3960
      Picture         =   "TestChg.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   1725
      TabIndex        =   2
      Top             =   3240
      Width           =   1725
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2475
      Index           =   0
      Left            =   2040
      Picture         =   "TestChg.frx":32E4
      ScaleHeight     =   2475
      ScaleWidth      =   1725
      TabIndex        =   1
      Top             =   3240
      Width           =   1725
   End
   Begin Project2.PictureChanger PictureChanger1 
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1725
      _extentx        =   3043
      _extenty        =   4366
   End
   Begin VB.Label Label1 
      Caption         =   "Steps"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Start using the copy change method.
    optChangeStyle(9).Value = True

    ' Start using 100 steps.
    txtSteps.Text = 100
End Sub

Private Sub optChangeStyle_Click(Index As Integer)
    PictureChanger1.ChangeStyle = Index
End Sub


' Copy the clicked picture into the PictureChanger.
Private Sub picImage_Click(Index As Integer)
    Set PictureChanger1.Picture = picImage(Index).Picture
End Sub


Private Sub txtSteps_Change()
    PictureChanger1.Steps = CInt(txtSteps.Text)
End Sub


