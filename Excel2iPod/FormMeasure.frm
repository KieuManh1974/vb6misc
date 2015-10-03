VERSION 5.00
Begin VB.Form FormMeasure 
   Caption         =   "Excel2iPod"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2625
   BeginProperty Font 
      Name            =   "Espy Sans"
      Size            =   9.75
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picText 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtSC 
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "."
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtCCW 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "7"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Spacer Character"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Column Character Width"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvert_Click()
    cmdConvert.Enabled = False
    ConvertExcel2iPod
    cmdConvert.Enabled = True
End Sub
