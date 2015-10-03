VERSION 5.00
Begin VB.Form frmPaper 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResponse 
      Alignment       =   1  'Right Justify
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
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblOperand 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblOperand 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAverage 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frmPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moNumberGenerator As clsNumberGenerator
Private moNumberRender As clsNumberRender
Private moResponses As clsStats

Private moCurrentQuestion As IGenerator

Private Sub Form_Activate()
    NextQuestion
End Sub

Private Sub Form_Initialize()
    Randomize
    Set moNumberGenerator = New clsNumberGenerator
    Set moNumberRender = New clsNumberRender
    Set moResponses = New clsStats
    
    'Set moResponses.Generator = moNumberGenerator
    moResponses.SetSize = 12
    Set moNumberRender.IRender_Paper = Me
    Randomize
End Sub

Private Sub NextQuestion()
    Dim vItem As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lCount As Long
    Dim dTime As Double
    
    Set moCurrentQuestion = moResponses.ChooseItem
    moNumberRender.IRender_Render moCurrentQuestion
    
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
        If moCurrentQuestion.CheckAnswer(Val(txtResponse.Text)) Then
            moResponses.AddResponse moCurrentQuestion, GetCounter
            NextQuestion
        Else
            txtResponse.ForeColor = vbRed
            KeyAscii = 0
        End If
    End If
End Sub
