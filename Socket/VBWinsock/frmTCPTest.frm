VERSION 5.00
Begin VB.Form frmTCPTest 
   Caption         =   "TCPIP"
   ClientHeight    =   4596
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   4896
   LinkTopic       =   "Form1"
   ScaleHeight     =   4596
   ScaleWidth      =   4896
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCloseConnection 
      Caption         =   "Close Conn"
      Height          =   372
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   2772
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4572
      Begin VB.TextBox txtReceive 
         Height          =   1812
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   720
         Width           =   3132
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Receive"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   972
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtSend 
         Height          =   372
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3132
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   372
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   372
      Left            =   1560
      TabIndex        =   5
      Text            =   "1001"
      Top             =   1080
      Width           =   1932
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Text            =   "100.0.1.2"
      Top             =   600
      Width           =   1932
   End
   Begin VB.TextBox txtLocalHost 
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Text            =   "100.0.1.1"
      Top             =   156
      Width           =   1932
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Remote Port:"
      Height          =   192
      Left            =   240
      TabIndex        =   4
      Top             =   1164
      Width           =   936
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Remote Host IP:"
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   684
      Width           =   1164
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local Host IP:"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   984
   End
End
Attribute VB_Name = "frmTCPTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tcp As VBWinsock.TCPIP

Private Sub cmdCloseConnection_Click()
    tcp.ShutdownConnection
    cmdSend.Enabled = False
    cmdReceive.Enabled = False
    Set tcp = Nothing
End Sub

Private Sub cmdConnect_Click()
    Set tcp = New VBWinsock.TCPIP
    With tcp
        .LocalHostIP = txtLocalHost.Text
        .RemoteHostIP = txtRemoteHost.Text
        .RemotePort = txtRemotePort.Text
        
        If Not .OpenConnection Then
            ShowError "Couldn't make connection"
        Else
            cmdSend.Enabled = True
            cmdReceive.Enabled = True
        End If
    End With
End Sub

Private Sub cmdReceive_Click()
    Dim strData As String
    Dim l As Long
    
    If tcp.IsDataAvailable Then
        If Not tcp.ReceiveData(strData, l) Then
            ShowError "Couldn't receive data"
        Else
            txtReceive.Text = strData
        End If
    Else
        MsgBox "Local Host hasn't received any data yet", vbInformation, "Information"
    End If
End Sub

Private Sub cmdSend_Click()
    If Not tcp.SendData(txtSend.Text & vbCrLf) Then ShowError "Couldn't send data"
End Sub

Private Sub ShowError(ByVal strMessage As String)
    MsgBox strMessage & vbCrLf & tcp.ErrorDescription, vbCritical, "Error"
End Sub
