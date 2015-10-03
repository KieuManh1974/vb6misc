VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate Rate"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   1935
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   1350
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   810
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   315
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "Calculate Exact Rate"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   90
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
    Dim d1 As Double
    Dim d2 As Double
    
    GetValues Text1.Text, d1, d2
    
    Text2.Text = d1
    Text3.Text = d2
End Sub

Public Sub GetValues(ByVal dRate As Double, dValue1 As Double, dValue2 As Double)
    Dim dRate1 As Double
    Dim dRate2 As Double
    Dim dDiff As Double
    Dim dThisDiff As Double
    
    dDiff = 1000
    For dRate1 = 0.0001 To 9.9999 Step 0.0001
        dRate1 = Int(dRate1 * 10000 + 0.5) / 10000
        dRate2 = Int((dRate / dRate1) * 10000 + 0.5) / 10000
        If Abs((dRate1 * dRate2) - dRate) < 0.000005 Then
            dThisDiff = Abs(dRate1 - dRate2)
            If dThisDiff < dDiff Then
                dDiff = dThisDiff
                dValue1 = dRate1
                dValue2 = dRate2
            End If
        End If
    Next
End Sub
