VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmFormPoster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Poster"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtURL 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Text            =   "http://www.response-o-matic.com/cgi-bin/rom.pl"
      Top             =   2950
      Width           =   3615
   End
   Begin VB.CommandButton CmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame FrameFields 
      Caption         =   "Fields"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin VB.TextBox TxtFields 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "FrmFormPoster.frx":0000
         Top             =   240
         Width           =   4335
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.response-o-matic.com"
      URL             =   "http://www.response-o-matic.com/cgi-bin/rom.pl"
      Document        =   "/cgi-bin/rom.pl"
   End
   Begin VB.Frame FrameUsernamePassword 
      Caption         =   "Username/ Password"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4575
      Begin VB.TextBox TxtPswd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Type your password here"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox TxtUname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   2
         Text            =   "Type your username here"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox ChkUP 
         Caption         =   "Username/ Password required"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label LblInfo 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label LblInfo 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1200
      End
   End
   Begin VB.Label LblURL 
      Caption         =   "Form URL:"
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "FrmFormPoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
   ' Make the application path the default directory
   ChDrive AppPath
   ChDir AppPath
   
   ' Centre the form on the screen
   With Me
      .Left = Screen.Width / 2 - .Width / 2
      .Top = Screen.Height / 2 - .Height / 2
   End With
   
   ' Load the form settings from the registry (if there)
   If (Len(GetSetting(App.EXEName, "Settings", "TxtFields", ""))) Then
      TxtFields = GetSetting(App.EXEName, "Settings", "TxtFields", "")
   End If
   
   If (Len(GetSetting(App.EXEName, "Settings", "TxtURL", ""))) Then
      TxtURL = GetSetting(App.EXEName, "Settings", "TxtURL", "")
   End If
   
   If (Len(GetSetting(App.EXEName, "Settings", "ChkUP", ""))) Then
      ChkUP = GetSetting(App.EXEName, "Settings", "ChkUP", "")
   End If
   
   If (Len(GetSetting(App.EXEName, "Settings", "TxtUname", ""))) Then
      TxtUname = GetSetting(App.EXEName, "Settings", "TxtUname", "")
   End If
   
   If (Len(GetSetting(App.EXEName, "Settings", "TxtUname", ""))) Then
      TxtPswd = GetSetting(App.EXEName, "Settings", "TxtPswd", "")
   End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' Save the form settings
   SaveSetting App.EXEName, "Settings", "TxtFields", TxtFields
   SaveSetting App.EXEName, "Settings", "TxtURL", TxtURL
   SaveSetting App.EXEName, "Settings", "ChkUP", ChkUP
   SaveSetting App.EXEName, "Settings", "TxtUname", TxtUname
   SaveSetting App.EXEName, "Settings", "TxtPswd", TxtPswd
End Sub

Private Sub CmdSubmit_Click()
   Dim Names() As String
   Dim Values() As String
   Dim Nvalues As Long
   Dim Lines() As String
   Dim I As Long
   
   ' Disable re- clicking
   CmdSubmit.Enabled = False
   
   ' Separate the lines in TxtFields into the Lines array
   Lines = Split(TxtFields, vbCrLf)
   Nvalues = 0
   For I = LBound(Lines) To UBound(Lines)
      If (InStr(Lines(I), ":") > 1) Then
         Nvalues = Nvalues + 1
         ReDim Preserve Names(1 To Nvalues)
         ReDim Preserve Values(1 To Nvalues)
         Names(Nvalues) = Mid$(Lines(I), 1, InStr(Lines(I), ":") - 1)
         Values(Nvalues) = Mid$(Lines(I), InStr(Lines(I), ":") + 1)
      End If
   Next
   If (Nvalues) Then
      Inet1.URL = TxtURL
      If (ChkUP.Value) Then
         ' Password required.
         Inet1.UserName = TxtUname
         Inet1.Password = TxtPswd
         
         ' For some reason (I don't know why) it seems necessary
         ' to insert the username & password into the URL of the form:
         ' http://username:Password@server.name.com/document.htp
         ' Otherwise it doesn't seem to work.
         ' if you find it fails by inserting these into the URL then
         ' remove the following lines.
         With Inet1
            Select Case .Protocol
               Case icHTTP, icDefault
                  .URL = "http://" & .UserName & ":" & .Password & "@" & _
                                         .RemoteHost & .Document
               Case icHTTPS ' don't know if this works, not tried it, don't see why not.
                  .URL = "https://" & .UserName & ":" & .Password & "@" & _
                                         .RemoteHost & .Document
               Case icFTP ' don't know if this works, not tried it, don't see why not.
                  .URL = "ftp://" & .UserName & ":" & .Password & "@" & _
                                         .RemoteHost & .Document
               Case Else
                  MsgBox "Unsupported protocol": Exit Sub
            End Select
         End With
      End If
   
      Inet1.Execute Inet1.URL, "POST", SendMessage(Names, Values)
   End If
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
   Dim vtData As Variant
   Dim strdata As String
   Dim Re As Integer
         
   Select Case State
      Case icError
         ' An error has occurred.
         MsgBox (Inet1.ResponseCode & ":" & Inet1.ResponseInfo)
         Inet1.Cancel
      Case icResponseCompleted
         strdata = ""
         ' Loop: get chunks of the response
         Do
            DoEvents
            vtData = Inet1.GetChunk(1024, icString)
            strdata = strdata & vtData
         Loop Until Len(vtData) = 0
         ' Save the response to "Return.htm" in the default directory
         Re = FreeFile
         Open AppPath & "Return.htm" For Output As #Re
         Print #Re, strdata
         Close (Re)
         
         ' Open "Return.htm" with the default browser
         Call ShellExecute(Me.hwnd, vbNullString, AppPath & "Return.htm", _
                                    vbNullString, vbNullString, SW_SHOWNORMAL)
         
         ' Can now re-enable CmdSubmit
         CmdSubmit.Enabled = True
   End Select
End Sub

Private Function SendMessage(Names() As String, Values() As String) As String
   
   Dim I As Long
   ' Concatenates the names and values arrays int the string data which
   ' is posted to the form
   SendMessage = ""
   For I = LBound(Names) To UBound(Names)
      SendMessage = SendMessage & Names(I) & "=" & Values(I)
      If (I <> UBound(Names)) Then SendMessage = SendMessage & "&"
   Next
End Function

Private Function AppPath() As String
   ' App.Path does not always have a backslash at the end.
   ' This function returns a string which definitely does.
   AppPath = App.Path
   If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\"
End Function
