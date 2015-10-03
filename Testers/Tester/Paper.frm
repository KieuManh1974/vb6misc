VERSION 5.00
Begin VB.Form Paper 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
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
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblAverage 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moBarGenerator As clsBarGenerator
Private moBarRenderer As clsRenderBars
Private moResponses As clsStats

Private mvCurrentQuestion As Variant

Private Sub Form_Activate()
    NextQuestion
End Sub

Private Sub Form_Initialize()
    Set moBarGenerator = New clsBarGenerator
    Set moBarRenderer = New clsRenderBars
    Set moResponses = New clsStats
    
    Set moResponses.Generator = moBarGenerator
    moResponses.SetSize = 12
    Set moBarRenderer.IRender_Paper = Me
    Randomize
End Sub

Private Sub NextQuestion()
    Dim vItem As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lCount As Long
    Dim dTime As Double
    
    mvCurrentQuestion = moResponses.ChooseItem
    moBarRenderer.IRender_Render (mvCurrentQuestion)
    
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
        If Val(txtResponse.Text) = mvCurrentQuestion Then
            moResponses.AddResponse mvCurrentQuestion, GetCounter
            NextQuestion
        Else
            txtResponse.ForeColor = vbRed
            KeyAscii = 0
        End If
    End If
End Sub
