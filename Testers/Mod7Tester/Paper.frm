VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Paper 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   142
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResponse 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAnswer 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   375
      Size            =   "661;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblNumber 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblAverage 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moQuestion As IQuestion
Private moResponses As clsStats

Private mvCurrentQuestion As Variant
Private Const msResponses As String = "senunikanakiku"

Private Sub cmdAnswer_Click(Index As Integer)
    Dim lIndex As Long
    
    For lIndex = 0 To 6
        cmdAnswer(lIndex).BackColor = &H8000000F
        cmdAnswer(lIndex).ForeColor = vbBlack
    Next
    
    If moQuestion.Compare(mvCurrentQuestion, Mid$(msResponses, Index * 2 + 1, 2)) Then
        moResponses.AddResponse mvCurrentQuestion, GetCounter
        NextQuestion
    Else
        cmdAnswer(Index).BackColor = vbRed
        cmdAnswer(Index).ForeColor = vbWhite
    End If
End Sub

Private Sub Form_Activate()
    NextQuestion
End Sub

Private Sub Form_Initialize()
    Set moQuestion = New clsDay
    Set moResponses = New clsStats
    
    Set moResponses.Generator = moQuestion
    Set moQuestion.Paper = Me
    Randomize
End Sub

Private Sub NextQuestion()
    Dim vItem As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lCount As Long
    Dim dTime As Double
    
    mvCurrentQuestion = moResponses.ChooseItem
    moQuestion.Render (mvCurrentQuestion)
    
    txtResponse.Text = ""
    StartCounter
    
    lblAverage.Caption = Format$(moResponses.Average, "0.00")

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
        If moQuestion.Compare(mvCurrentQuestion, txtResponse.Text) Then
            moResponses.AddResponse mvCurrentQuestion, GetCounter
            NextQuestion
        Else
            txtResponse.ForeColor = vbRed
            KeyAscii = 0
        End If
    End If
End Sub
