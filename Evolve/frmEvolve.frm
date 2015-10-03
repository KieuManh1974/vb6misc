VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1680
   End
   Begin VB.CommandButton cmdMutate 
      Caption         =   "M"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moPop As New clsPopulation

Private mlLastSelection As Long

Private Sub cmdMutate_Click()
    moPop.MutateAll
    Me.Cls
    moPop.Render
End Sub

Private Sub Form_Load()
    Randomize
    glCellWidth = 100
    moPop.CanvasDC = Me.hDC
    moPop.Render
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lColumn As Long
    Dim lRow As Long
    
    lColumn = (X - 50) \ glCellWidth
    lRow = (Y - 50) \ glCellWidth
    If lColumn < 0 Or lColumn > 2 Or lRow < 0 Or lRow > 2 Then
        Exit Sub
    End If
    mlLastSelection = lColumn + lRow * 3
    moPop.SelectAndMutate lColumn + lRow * 3
    Me.Cls
    moPop.Render
End Sub

Private Sub Timer1_Timer()
    moPop.SelectAndMutate mlLastSelection
    Me.Cls
    moPop.Render
End Sub
