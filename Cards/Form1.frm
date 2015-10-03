VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHand 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblProbability 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    InitialiseParser
End Sub

Private Sub txtHand_Change()
    Dim oResult As New ParseTree
    
    Stream.Text = txtHand.Text
    
    If oParser.Parse(oResult) Then
        lblProbability.Caption = Probability(oResult)
    Else
        lblProbability.Caption = "0"
    End If
End Sub

Private Function Probability(oResult As ParseTree) As String
    Dim lDeck As Long
    Dim oCard As ParseTree
    Dim lNumerator As Long
    Dim lDenominator As Long
    Dim lNumberCount(1 To 13) As Long
    Dim lSuitCount(1 To 4) As Long
    Dim lCardCount(1 To 52) As Long
    Dim lIndex As Long
    Dim sNumbers As String
    Dim sSuits As String
    Dim lCard As Long
    
    sNumbers = "A234567890JQK"
    sSuits = "DCHS"
    
    For lIndex = 1 To 13
        lNumberCount(lIndex) = 4
    Next
    
    For lIndex = 1 To 4
        lSuitCount(lIndex) = 13
    Next
    
    For lIndex = 1 To 52
        lCardCount(lIndex) = 1
    Next
    
    lDeck = 52
    lNumerator = 1
    lDenominator = 1
    
    For Each oCard In oResult.SubTree
        Select Case oCard(1).Index
            Case 1
                lCard = (InStr(sNumbers, oCard(1)(1)(1).Text) - 1 + (InStr(sSuits, UCase$(oCard(1)(1)(2).Text)) - 1) * 13) + 1
                lNumerator = lNumerator * lCardCount(lCard)
                lCardCount(lCard) = lCardCount(lCard) - 1
                lDenominator = lDenominator * lDeck
                lDeck = lDeck - 1
            Case 2 'num
                lCard = InStr(sNumbers, UCase$(oCard.Text))
                lNumerator = lNumerator * lNumberCount(lCard)
                lNumberCount(lCard) = lNumberCount(lCard) - 1
                lDenominator = lDenominator * lDeck
                lDeck = lDeck - 1
            Case 3 'suit
                lCard = InStr(sSuits, UCase$(oCard.Text))
                lNumerator = lNumerator * lSuitCount(lCard)
                lSuitCount(lCard) = lSuitCount(lCard) - 1
                lDenominator = lDenominator * lDeck
                lDeck = lDeck - 1
        End Select
    Next
    
    Probability = lNumerator & " in " & lDenominator
End Function
