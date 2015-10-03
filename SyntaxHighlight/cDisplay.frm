VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cDisplay 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox moRTF 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"cDisplay.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "cDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moHinting As New cHinting

Private Sub Form_Load()
    moHinting.HintText "a"
End Sub

Private Sub Form_Resize()
    moRTF.Width = Me.Width
    moRTF.Height = Me.Height
End Sub

Private Sub moRTF_Change()
'    Dim lPos As Long
'    Dim lLength As Long
'
'    lPos = moRTF.SelStart
'    lLength = moRTF.SelLength
'
'    moRTF.TextRTF = moHinting.HintText(moRTF.Text)
'    moRTF.SelStart = lPos
'    moRTF.SelLength = lLength

Debug.Print moRTF.SelStart
End Sub
