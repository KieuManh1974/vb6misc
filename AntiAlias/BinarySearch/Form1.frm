VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Binary 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtValue 
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblMultipliedInt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin MSForms.CommandButton cmdParent 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
      BackColor       =   16777215
      Caption         =   "Parent"
      Size            =   "2990;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdHigher 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
      BackColor       =   16777215
      Caption         =   "Higher"
      Size            =   "2990;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdLower 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      BackColor       =   16777215
      Caption         =   "Lower"
      Size            =   "2990;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblMultiplied 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblBinary 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "Binary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msBinary As String
Private mlValue As Long

Private Sub cmdHigher_Click()
    msBinary = msBinary & "1"
    UpdateBinary
End Sub

Private Sub cmdLower_Click()
    msBinary = msBinary & "0"
    UpdateBinary
End Sub

Private Sub cmdParent_Click()
    If Len(msBinary) > 1 Then
        msBinary = Left(msBinary, Len(msBinary) - 1)
        UpdateBinary
    End If
End Sub

Private Sub Form_Load()
    msBinary = "1"
    UpdateBinary
End Sub

Private Sub txtValue_Change()
    mlValue = Val(txtValue)
    UpdateBinary
End Sub

Private Sub txtValue_LostFocus()
    mlValue = Val(txtValue)
    UpdateBinary
End Sub

Private Function Base(ByVal lnum As String, ByVal frombase As Integer, ByVal tobase As Integer) As String
    Const digits As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim digitpos As Long
    Dim power As Long
    Dim result As Long
    Dim digitvalue As Long
    
    power = 1
    For digitpos = Len(lnum) To 1 Step -1
        digitvalue = InStr(digits, Mid$(lnum, digitpos, 1)) - 1
        result = result + digitvalue * power
        power = power * frombase
    Next
    
    power = tobase
    While result > 0
        digitvalue = ((result / power) - Int(result / power)) * power
        Base = Mid$(digits, digitvalue + 1, 1) & Base
        result = Int(result / power)
    Wend
End Function

Private Sub UpdateBinary()
    Dim dValue As Double
    
    dValue = mlValue * (Base(msBinary, 2, 10) * 2 - (2 ^ Len(msBinary)) + 1) / (2 ^ Len(msBinary))
    
    lblMultiplied.Caption = dValue
    lblMultipliedInt.Caption = Int(dValue + 0.5)
    
    lblBinary.Caption = msBinary
    
End Sub
