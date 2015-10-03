VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Curry Menu"
   ClientHeight    =   4728
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   7488
   LinkTopic       =   "Form1"
   ScaleHeight     =   4728
   ScaleWidth      =   7488
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   11
      Left            =   2760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   10
      Left            =   1920
      Picture         =   "Form1.frx":215BF2A
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   9
      Left            =   1053
      Picture         =   "Form1.frx":42564EC
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   9
      Top             =   2106
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   8
      Left            =   234
      Picture         =   "Form1.frx":42B0536
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   8
      Top             =   2106
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   7
      Left            =   5967
      Picture         =   "Form1.frx":430A580
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   7
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   6
      Left            =   5148
      Picture         =   "Form1.frx":43645CA
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   6
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   5
      Left            =   4329
      Picture         =   "Form1.frx":43BE614
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   5
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   4
      Left            =   3510
      Picture         =   "Form1.frx":441865E
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   4
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   3
      Left            =   2691
      Picture         =   "Form1.frx":44726A8
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   3
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   2
      Left            =   1872
      Picture         =   "Form1.frx":44CC6F2
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   2
      Top             =   1404
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   1
      Left            =   1080
      Picture         =   "Form1.frx":452673C
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.PictureBox PIC 
      Height          =   598
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":4580786
      ScaleHeight     =   552
      ScaleWidth      =   672
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.Image Image1 
      Height          =   996
      Left            =   0
      Picture         =   "Form1.frx":45DA7D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   996
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const TotalNumber = 12

Private Sub Form_Load()
    Image1_MouseUp 0, 0, 0, 0
End Sub

Private Sub Form_Resize()
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static state As Long


    If Button = vbLeftButton Then
        state = (state + 1) Mod TotalNumber
    ElseIf Button = vbRightButton Then
        state = (TotalNumber + state - 1) Mod TotalNumber
    End If

    Image1.Picture = PIC(state).Picture
End Sub
