VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Operation"
      Height          =   615
      Left            =   45
      TabIndex        =   26
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optOperation 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optOperation 
         Caption         =   "Subtract"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Spell"
      Height          =   615
      Left            =   45
      TabIndex        =   22
      Top             =   4920
      Width           =   2295
      Begin VB.OptionButton optSpell 
         Caption         =   "Full"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optSpell 
         Caption         =   "Digit"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optSpell 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Configuration"
      Height          =   615
      Left            =   45
      TabIndex        =   19
      Top             =   4320
      Width           =   2295
      Begin VB.OptionButton optConfiguration 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optConfiguration 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Carry/Borrow Overflow"
      Height          =   615
      Left            =   45
      TabIndex        =   14
      Top             =   3720
      Width           =   2295
      Begin VB.OptionButton optCarryBorrowOverflow 
         Caption         =   "Disallow"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCarryBorrowOverflow 
         Caption         =   "Allow"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Carry/Borrow In"
      Height          =   615
      Left            =   45
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
      Begin VB.OptionButton optCarryBorrowIn 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarryBorrowIn 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarryBorrowIn 
         Caption         =   "Yes"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Number of Digits"
      Height          =   615
      Left            =   45
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      Begin VB.TextBox txtNumDigits 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Carry/Borrow Out"
      Height          =   975
      Left            =   45
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
      Begin VB.Frame Frame4 
         Caption         =   "Carry/Borrow Out"
         Height          =   975
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   2775
         Begin VB.TextBox txtCarryBorrowPatternOut 
            Height          =   285
            Left            =   840
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optCarryBorrowOut 
            Caption         =   "Yes"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   33
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton optCarryBorrowOut 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton optCarryBorrowOut 
            Caption         =   "Any"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optCarryBorrowOut 
            Caption         =   "No Double"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.OptionButton optCarryOut 
         Caption         =   "No Double"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optCarryOut 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCarryOut 
         Caption         =   "No"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optCarryOut 
         Caption         =   "Yes"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parity"
      Height          =   975
      Left            =   45
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
      Begin VB.OptionButton optParity 
         Caption         =   "Any"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Even"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Odd"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optParity 
         Caption         =   "Odd / Even"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   4
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
    txtNumDigits = glNumDigits
    
    optOperation(glOperation + 1).Value = True
    optParity(glParity + 1).Value = True
    optCarryBorrowOut(glCarryBorrowOut + 1).Value = True
    optCarryBorrowIn(glCarryBorrowIn + 1).Value = True
    optCarryBorrowOverflow(glCarryBorrowOverflow).Value = True
    optConfiguration(glConfiguration).Value = True
    optSpell(glSpell).Value = True
End Sub

Private Sub optCarryBorrowOverflow_Click(Index As Integer)
    glCarryBorrowOverflow = Index
End Sub

Private Sub optCarryBorrowIn_Click(Index As Integer)
    glCarryBorrowIn = Index - 1
End Sub

Private Sub optCarryBorrowOut_Click(Index As Integer)
    glCarryBorrowOut = Index - 1
    SetCarryPattern
End Sub

Private Sub optConfiguration_Click(Index As Integer)
    glConfiguration = Index
End Sub

Private Sub optOperation_Click(Index As Integer)
    glOperation = Index - 1
End Sub

Private Sub optParity_Click(Index As Integer)
    glParity = Index - 1
End Sub

Private Sub optSpell_Click(Index As Integer)
    glSpell = Index
End Sub

Private Sub txtCarryBorrowPatternOut_Change()
    glCarryBorrowPatternOut = txtCarryBorrowPatternOut
End Sub


Private Sub txtNumDigits_Change()
    glNumDigits = Val(txtNumDigits.Text)
    SetCarryPattern
End Sub

Private Function SetCarryPattern()
    Select Case glCarryBorrowOut
        Case 0
            txtCarryBorrowPatternOut.Text = String$(glNumDigits, "0")
        Case 1
            txtCarryBorrowPatternOut.Text = String$(glNumDigits, "1")
        Case Else
            txtCarryBorrowPatternOut.Text = glCarryBorrowPatternOut
    End Select
End Function
