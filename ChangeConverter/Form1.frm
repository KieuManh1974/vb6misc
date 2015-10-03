VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemainder 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      TabIndex        =   50
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   8
      Left            =   2520
      TabIndex        =   49
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   8
      Left            =   2520
      TabIndex        =   47
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   45
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   43
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   39
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   37
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   35
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   7
      Left            =   3120
      TabIndex        =   33
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   31
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   29
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   11
      Left            =   720
      TabIndex        =   27
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   23
      Text            =   "21"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   21
      Text            =   "5"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   19
      Text            =   "19"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   17
      Text            =   "12"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   15
      Text            =   "9"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   13
      Text            =   "1"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   7
      Left            =   3120
      TabIndex        =   11
      Text            =   "3"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   7
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   11
      Left            =   720
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtFloat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "Remainder"
      Height          =   255
      Left            =   1200
      TabIndex        =   51
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "£2"
      Height          =   255
      Left            =   2520
      TabIndex        =   48
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label24 
      Caption         =   "£2"
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label23 
      Caption         =   "1p"
      Height          =   255
      Left            =   6720
      TabIndex        =   44
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "2p"
      Height          =   255
      Left            =   6120
      TabIndex        =   42
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "5p"
      Height          =   255
      Left            =   5520
      TabIndex        =   40
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label20 
      Caption         =   "10p"
      Height          =   255
      Left            =   4920
      TabIndex        =   38
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "20p"
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "50p"
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "£1"
      Height          =   255
      Left            =   3120
      TabIndex        =   32
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "£5"
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "£10"
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "£20"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "£50"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "1p"
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "2p"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "5p"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "10p"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "20p"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "50p"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "£1"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "£5"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "£10"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "£20"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "£50"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.Text) Then
        PopulateChange
    End If
End Sub

Private Sub PopulateChange()
    Dim cAmount As Currency
    Dim lCashIndex As Long
    Dim cCashAmount As Currency
    Dim lNumber As Long
    Dim vAmounts As Variant
    
    vAmounts = Array(0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50)
    
    lCashIndex = 12
    cAmount = Val(txtAmount.Text)
    
    While lCashIndex > 0
        cCashAmount = vAmounts(lCashIndex - 1)
        lNumber = Int(cAmount / cCashAmount)
        If txtFloat(lCashIndex).Text <> "" Then
            If IsNumeric(txtFloat(lCashIndex).Text) Then
                If lNumber > Int(Val(txtFloat(lCashIndex).Text)) Then
                    lNumber = Int(Val(txtFloat(lCashIndex).Text))
                End If
            End If
        End If
        
        txtChange(lCashIndex).Text = lNumber
        cAmount = cAmount - lNumber * cCashAmount
        
        lCashIndex = lCashIndex - 1
    Wend
    txtRemainder.Text = cAmount
End Sub

Private Sub txtFloat_Change(Index As Integer)
    PopulateChange
End Sub
