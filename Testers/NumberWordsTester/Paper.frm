VERSION 5.00
Begin VB.Form Paper 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResponse 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblNumber 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblAverage 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moNumberQuestion As clsNumberQuestion
Private moResponses As clsStats

Private mvCurrentQuestion As Variant

Private Sub Form_Activate()
    NextQuestion
End Sub

Private Sub Form_Initialize()
    Set moNumberQuestion = New clsNumberQuestion
    Set moResponses = New clsStats
    
    Set moResponses.Generator = moNumberQuestion
    Set moNumberQuestion.IQuestion_Paper = Me
    Randomize
End Sub

Private Sub NextQuestion()
    Dim vItem As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lCount As Long
    Dim dTime As Double
    
    mvCurrentQuestion = moResponses.ChooseItem
    moNumberQuestion.IQuestion_Render (mvCurrentQuestion)
    
    txtResponse.Text = ""
    StartCounter
    
'    lblAverage.Caption = Format$(moResponses.Average, "0.00")
'
'    Me.CurrentX = 200
'    Me.CurrentY = 0
'    For lIndex = 1 To 12
'        lCount = 0
'        dTime = 0
'        For lIndex2 = 1 To moResponses.moResponses.Count
'            If moResponses.moResponses(lIndex2).Item = lIndex Then
'                lCount = moResponses.moResponses(lIndex2).Count
'                dTime = moResponses.moResponses(lIndex2).CumulativeResponseTime / moResponses.moResponses(lIndex2).Count
'                Exit For
'            End If
'        Next
'        Me.CurrentX = 200
'        Me.Print lIndex & ":" & lCount & " " & Format$(dTime, "0.00")
'    Next
End Sub


Private Sub txtResponse_Change()
    txtResponse.ForeColor = vbBlack
End Sub

Private Sub txtResponse_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If moNumberQuestion.IQuestion_Compare(mvCurrentQuestion, txtResponse.Text) Then
            moResponses.AddResponse mvCurrentQuestion, GetCounter
            NextQuestion
        Else
            txtResponse.ForeColor = vbRed
            KeyAscii = 0
        End If
    End If
End Sub
