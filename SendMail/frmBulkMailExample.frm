VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bulk Mail Demonstration Client for vbSendMail Component"
   ClientHeight    =   6165
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7455
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   360
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   2055
      Left            =   6060
      TabIndex        =   23
      Top             =   1440
      Width           =   1275
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Text            =   "1"
         Top             =   900
         Width           =   555
      End
      Begin VB.OptionButton optSend 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Use the Send Method only for each message."
         Top             =   1380
         Value           =   -1  'True
         Width           =   1075
      End
      Begin VB.OptionButton optSend 
         Caption         =   "Bulk Send"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Use the Connect, Send and Disconnect Methods."
         Top             =   1680
         Width           =   1075
      End
      Begin VB.Label lblQty 
         Caption         =   "Number of Messages to send:"
         Height          =   555
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   6060
      TabIndex        =   22
      Top             =   900
      Width           =   1275
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   1680
      TabIndex        =   20
      Top             =   4320
      Width           =   4200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   6060
      TabIndex        =   8
      Top             =   3960
      Width           =   1275
   End
   Begin VB.TextBox txtAttach 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   4200
   End
   Begin VB.TextBox txtMsg 
      Height          =   1680
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2220
      Width           =   4200
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1860
      Width           =   4200
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   780
      Width           =   4200
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   420
      Width           =   4200
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1500
      Width           =   4200
   End
   Begin VB.TextBox txtToName 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1140
      Width           =   4200
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   75
      Width           =   4200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   6060
      TabIndex        =   10
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   6060
      TabIndex        =   9
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   28
      Top             =   5820
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   4380
      Width           =   555
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   5580
      Width           =   750
   End
   Begin VB.Label lblAttach 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label lblFrom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label lblFromName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblToName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   105
      Width           =   1140
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *****************************************************************************
' This sample application is a simple demonstration of sending bulk mail. It
' sends a mail message to the SAME recipient a selectable number of times. Its
' purpose is simply to demonstrate the two methods available for sending bulk
' mail and the performance differences between them.
'
' In a real application you would need to load a recipient list from a file
' or database.
' *****************************************************************************

Option Explicit
Option Compare Text

Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
Private bSendFailed     As Boolean


Private Sub cmdSend_Click()

    Dim lCount      As Long
    Dim lCtr        As Long
    Dim t!

    cmdSend.Enabled = False
    bSendFailed = False
    lstStatus.Clear
    lblTime.Caption = ""
    Screen.MousePointer = vbHourglass

    With poSendMail

        ' **************************************************************************
        ' Set the basic properties common to all messages to be sent
        ' **************************************************************************
        .SMTPHost = txtServer.Text                  ' Required the fist time, optional thereafter
        .From = txtFrom.Text                        ' Required the fist time, optional thereafter
        .FromDisplayName = txtFromName.Text         ' Optional, saved after first use
        .Message = txtMsg.Text                      ' Optional
        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

        ' get the message count and set the timer
        lCount = Val(txtQty)
        If lCount = 0 Then Exit Sub
        t! = Timer

        ' **************************************************************************
        ' Send the mail in a loop. In a real app you would need to load a new
        ' recipient from a file or database each pass through the loop.
        ' **************************************************************************
        If optSend(0).Value = True Then

            ' send method only (normal button)
            For lCtr = 1 To lCount
                .Recipient = txtTo.Text
                .RecipientDisplayName = txtToName.Text
                .Subject = txtSubject & " (Message # " & Str(lCtr) & ")"
                lblTime = "Sending message " & Str(lCtr)
                .Send
            Next

        Else
            ' connect, send, & disconnect methods (bulk send button)
            If .Connect Then
                For lCtr = 1 To lCount
                    lblTime = "Sending message " & Str(lCtr)
                    .Recipient = txtTo.Text
                    .RecipientDisplayName = txtToName.Text
                    .Subject = txtSubject & " (Message # " & Str(lCtr) & ")"
                    .Send
                Next
                .Disconnect
            End If
        End If

    End With

    ' display the results
    If Not bSendFailed Then lblTime.Caption = Str(lCount) & " Messages sent in " & Format$(Timer - t!, "#,##0.0") & " seconds."
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event'

    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    bSendFailed = True
    lblProgress = ""
    lblTime = ""

End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'

    lblProgress = "Send Successful!"

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'

    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()

    Set poSendMail = New clsSendMail

    With Me
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        .lblProgress = ""
        .lblTime = ""
        .Show
        .Refresh
    End With

    With poSendMail
        .SMTPHostValidation = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
    End With

    RetrieveSavedValues

    cmDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set poSendMail = Nothing

End Sub

Private Sub RetrieveSavedValues()

' *****************************************************************************
' Retrieve saved values by reading the components 'Persistent' properties
' *****************************************************************************

    txtServer.Text = poSendMail.SMTPHost
    txtFrom.Text = poSendMail.From
    txtFromName.Text = poSendMail.FromDisplayName

End Sub

Private Sub cmdBrowse_Click()

    cmDialog.ShowOpen

    If txtAttach.Text = "" Then
        txtAttach.Text = cmDialog.FileName
    Else
        txtAttach.Text = txtAttach.Text & ";" & cmDialog.FileName
    End If

End Sub

Private Sub cmdExit_Click()

Dim frm As Form

For Each frm In Forms
    Unload frm
    Set frm = Nothing
Next

End

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    lblTime = ""
    RetrieveSavedValues

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 48 To 57                               ' numeric
        Case 8                                      ' backspace
        Case Else: KeyAscii = 0
    End Select

End Sub

