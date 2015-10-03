VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   14505
   ClientTop       =   0
   ClientWidth     =   840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   840
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   840
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "ABC"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iCount As Integer

Private Sub Form_Load()
    iCount = Val(Inet1.OpenURL("http://www.marsman.demon.co.uk/count.txt"))
    Label1.Caption = iCount
    Me.Caption = iCount
End Sub

Private Sub Timer1_Timer()
    Dim iThisCount As Integer
    
    iThisCount = Val(Inet1.OpenURL("http://www.marsman.demon.co.uk/count.txt"))
    Label1.Caption = iThisCount
    Me.Caption = iThisCount
    If iThisCount <> iCount Then
        Beep
        iCount = iThisCount
    End If
End Sub
