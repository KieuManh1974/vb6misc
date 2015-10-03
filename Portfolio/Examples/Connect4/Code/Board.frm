VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Board 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect4"
   ClientHeight    =   3750
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtOpponent 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "BR005151"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2630
      Left            =   180
      ScaleHeight     =   2650
      ScaleMode       =   0  'User
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   165
      Width           =   2650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Opponent Machine:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private columnheight(0 To 7) As Integer
Private Board(-3 To 7 + 3, -3 To 7 + 3) As Long
Private myToggle As Boolean

Private Property Let Toggle(bToggle As Boolean)
    myToggle = bToggle
    If bToggle Then
        Me.Caption = "Connect4 - My Move"
    Else
        Me.Caption = "Connect4 - Your Move"
    End If
End Property

Private Property Get Toggle() As Boolean
    Toggle = myToggle
End Property

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Toggle = False
    txtOpponent.Text = GetSetting("Connect4", "Opponent", "MachineId", "BR005151")
    Winsock1.RemoteHost = txtOpponent.Text
    Winsock1.RemotePort = 1002
    Winsock1.Bind 1001
    
    DrawBoard
End Sub
   
Private Sub Command1_Click()
    On Error Resume Next
    Winsock1.SendData CInt(-1)
    DrawBoard
    Picture1.Enabled = True
    Toggle = False
End Sub

Private Sub DrawBoard()
    Dim xline As Integer
    Dim xpos As Integer
    Dim ypos As Integer
    
    Picture1.Cls
    For xline = 0 To 7
        Picture1.Line (0, xline * 375)-Step(7 * 375, 0), vbBlack
        Picture1.Line (xline * 375, 0)-Step(0, 7 * 375), vbBlack
    Next
    
    Picture1.FillColor = Picture1.BackColor
    For xpos = 0 To 6
        For ypos = 0 To 6
            Picture1.Circle (xpos * 375 + 375 \ 2, ypos * 375 + 375 \ 2), 150, vbBlack
            Board(xpos, ypos) = 0
            columnheight(xpos) = 0
        Next
    Next
    
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show vbModal
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Toggle Then
        Exit Sub
    End If
    If Button = vbLeftButton Then
        PlacePiece X \ 375, vbRed
        Winsock1.SendData CInt(X \ 375)
        Toggle = True
    Else
        'PlacePiece X \ 375, vbBlue
    End If
End Sub

Private Sub txtOpponent_Change()
    SaveSetting "Connect4", "Opponent", "MachineId", txtOpponent.Text
End Sub

Private Sub txtOpponent_LostFocus()
    Winsock1.RemoteHost = txtOpponent.Text
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim xindex As Integer
    Winsock1.GetData xindex, vbInteger, bytesTotal
    Winsock1.RemoteHost = Winsock1.RemoteHostIP
    txtOpponent.Text = Winsock1.RemoteHostIP
    If xindex = -1 Then
        DrawBoard
        Picture1.Enabled = True
    Else
        Toggle = False
        PlacePiece xindex, vbBlue
    End If
End Sub

Private Sub PlacePiece(ByVal xindex As Integer, ByVal piececolour As Long)

    If xindex > 6 Then Exit Sub
    
    Picture1.FillColor = piececolour
    Picture1.FillStyle = 0
    Dim xcoord As Single
    Dim ycoord As Single
    Dim radius As Single
    
    If columnheight(xindex) = 7 Then
        Exit Sub
    End If
    xcoord = xindex * 375 + 375 \ 2
    ycoord = (6 - columnheight(xindex)) * 375 + 375 \ 2
    radius = 150
    Picture1.Circle (xcoord, ycoord), radius, vbBlack
    
    Board(xindex, 6 - columnheight(xindex)) = piececolour
    columnheight(xindex) = columnheight(xindex) + 1

    ' Scan board
    Dim ypos As Integer
    Dim xpos As Integer
    Dim xdir As Integer
    Dim ydir As Integer
    Dim mcount As Integer
    Dim previouspiece  As Long
    Dim z As Integer
    
    Dim directions As Variant
    Dim directionindex As Integer
    
    directions = Array(Array(1, -1), Array(1, 0), Array(1, 1), Array(0, 1))
    
    For directionindex = 0 To UBound(directions)
        xdir = directions(directionindex)(0)
        ydir = directions(directionindex)(1)
        For ypos = 0 To 6
            For xpos = 0 To 6
                mcount = 0
                previouspiece = 0
                For z = 0 To 3
                    If (Board(xpos + z * xdir, ypos + z * ydir) = previouspiece) And (previouspiece <> 0) Then
                        mcount = mcount + 1
                    End If
                    previouspiece = Board(xpos + z * xdir, ypos + z * ydir)
                    If mcount = 3 Then
                        Picture1.DrawWidth = 5
                        Picture1.Line (xpos * 375 + 375 \ 2, ypos * 375 + 375 \ 2)-Step(xdir * 375 * 3, ydir * 375 * 3), vbWhite
                        Picture1.DrawWidth = 1
                        Picture1.Enabled = False
                        Exit Sub
                    End If
                Next
            Next
        Next
    Next
End Sub

