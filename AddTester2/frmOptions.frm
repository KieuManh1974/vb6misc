VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "Carry Pattern"
      Height          =   615
      Left            =   50
      TabIndex        =   28
      Top             =   5040
      Width           =   2295
      Begin VB.TextBox txtCarryPattern 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Carry Overflow"
      Height          =   615
      Left            =   50
      TabIndex        =   23
      Top             =   3240
      Width           =   2295
      Begin VB.OptionButton optCarryOverflow 
         Caption         =   "Disallow"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCarryOverflow 
         Caption         =   "Allow"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Carry Propagation"
      Height          =   615
      Left            =   50
      TabIndex        =   20
      Top             =   2640
      Width           =   2295
      Begin VB.OptionButton optCarryPropagation 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarryPropagation 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarryPropagation 
         Caption         =   "Yes"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Number of Digits"
      Height          =   615
      Left            =   50
      TabIndex        =   4
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtNumDigits 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Distance"
      Height          =   615
      Left            =   50
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
      Begin VB.OptionButton optDistance 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optDistance 
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDistance 
         Caption         =   "1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDistance 
         Caption         =   "2"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Centres"
      Height          =   615
      Left            =   50
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
      Begin VB.OptionButton optCentres 
         Caption         =   "Yes"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCentres 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCentres 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Carry"
      Height          =   1095
      Left            =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
      Begin VB.OptionButton optCarry 
         Caption         =   "No Double"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCarry 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarry 
         Caption         =   "No"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optCarry 
         Caption         =   "Yes"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parity"
      Height          =   975
      Left            =   50
      TabIndex        =   0
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton optParity 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Even"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Odd"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Odd / Even"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtNumDigits = glNumdigits
    
    optParity(glParity + 1).Value = True
    optCarry(glCarry + 1).Value = True
    optDistance(glDistance + 1).Value = True
    optCentres(glCentres + 1).Value = True
    optCarryPropagation(glCarryPropagation + 1).Value = True
    optCarryOverflow(glCarryOverflow).Value = True
End Sub

Private Sub optCarry_Click(Index As Integer)
    glCarry = Index - 1
    SetCarryPattern
End Sub

Private Sub optCarryOverflow_Click(Index As Integer)
    glCarryOverflow = Index
End Sub

Private Sub optCarryPropagation_Click(Index As Integer)
    glCarryPropagation = Index - 1
End Sub

Private Sub optDistance_Click(Index As Integer)
    glDistance = Index - 1
End Sub

Private Sub optParity_Click(Index As Integer)
    glParity = Index - 1
End Sub

Private Sub txtCarryPattern_Change()
    glCarryPattern = txtCarryPattern
End Sub

Private Sub txtNumDigits_Change()
    glNumdigits = Val(txtNumDigits.Text)
    SetCarryPattern
End Sub

Private Function SetCarryPattern()
    Select Case glCarry
        Case 1
            txtCarryPattern.Text = String$(glNumdigits, "0")
        Case 2
            txtCarryPattern.Text = String$(glNumdigits, "1")
        Case Else
            txtCarryPattern.Text = glCarryPattern
    End Select
End Function
