VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1200
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "10.10.13.104"
      RemotePort      =   7
      LocalPort       =   7
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10.10.2.44

Option Explicit

Private lIntervalMinutes As Long

Private Sub Form_Load()
    Winsock1.Close
    Winsock1.Bind 7
    'Winsock1.SendData "TEST"
    'Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Winsock1.SendData "TEST"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim vPacket As Variant

    Winsock1.GetData vPacket, vbString
End Sub

