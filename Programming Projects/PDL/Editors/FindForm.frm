VERSION 5.00
Begin VB.Form FindForm 
   Caption         =   "Find & Replace"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "Whole Document"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Selected Text"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FindForm.frx":0000
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

