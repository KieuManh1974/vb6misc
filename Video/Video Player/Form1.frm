VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WMPLibCtl.WindowsMediaPlayer wm 
      CausesValidation=   0   'False
      Height          =   11520
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   15360
      URL             =   "D:\Media\Video\Wildlife on 2\Wildlife on 2 - The Butteryfly Beauty Or The Beast.avi"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   27093
      _cy             =   20320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iNotPlay As Boolean

Private Sub Form_Activate()
    Dim lTVTool As Long
    Dim iLastPosition As Long
    
    If Command <> "" Then
        If Left$(Command, 1) = """" Then
            wm.URL = Mid$(Command, 2, Len(Command) - 2)
        Else
            wm.URL = Command
        End If
    End If

    iLastPosition = GetSetting("VideoPlayer", "LastPosition", "p", 1)
    wm.Left = GetSetting("VideoPlayer", "Position" & iLastPosition, "x", wm.Left)
    wm.Top = GetSetting("VideoPlayer", "Position" & iLastPosition, "y", wm.Top)
    wm.Width = GetSetting("VideoPlayer", "Position" & iLastPosition, "w", wm.Width)
    
    lTVTool = OpenApplication("D:\Program Files\TVTool 8\TVTOOL.exe", "D:\Program Files\TVTool 8\")
    Wait 5, lTVTool
    CloseApplication lTVTool
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iIndex As Integer
    
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF
            wm.Height = wm.Height + 1
        Case vbKeyV
            wm.Height = wm.Height - 1
        Case vbKeyX
            Select Case Shift
                Case 0
                    wm.Left = wm.Left + 1
                Case 1
                    Me.Left = Me.Left + 150
            End Select
        Case vbKeyZ
            Select Case Shift
                Case 0
                    wm.Left = wm.Left - 1
                Case 1
                    Me.Left = Me.Left - 150
            End Select
        Case vbKeyC
            Select Case Shift
                Case 0
                    wm.Top = wm.Top + 1
                Case 1
                    Me.Top = Me.Top + 150
            End Select
        Case vbKeyD
            Select Case Shift
                Case 0
                    wm.Top = wm.Top - 1
                Case 1
                    Me.Top = Me.Top - 150
            End Select
        Case vbKeyA
            wm.Width = wm.Width - 2
            wm.Left = wm.Left + 1
        Case vbKeyS
            wm.Width = wm.Width + 2
            wm.Left = wm.Left - 1
        Case vbKeyP, vbKeyQ, vbKeyW, vbKeyE, vbKeyR, vbKeyT, vbKeyY, vbKeyU, vbKeyI, vbKeyO
            iIndex = InStr("PQWERTYUIO", Chr$(KeyCode)) - 1
            SaveSetting "VideoPlayer", "Position" & iIndex, "x", wm.Left
            SaveSetting "VideoPlayer", "Position" & iIndex, "y", wm.Top
            SaveSetting "VideoPlayer", "Position" & iIndex, "w", wm.Width
            SaveSetting "VideoPlayer", "LastPosition", "p", iIndex
        Case 48 To 57
            wm.Left = GetSetting("VideoPlayer", "Position" & KeyCode - 48, "x", wm.Left)
            wm.Top = GetSetting("VideoPlayer", "Position" & KeyCode - 48, "y", wm.Top)
            wm.Width = GetSetting("VideoPlayer", "Position" & KeyCode - 48, "w", wm.Width)
            SaveSetting "VideoPlayer", "LastPosition", "p", KeyCode - 48
        Case 27
            End
        Case vbKeyM
            Me.WindowState = vbMaximized
        Case vbKeyN
            Me.WindowState = vbNormal
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
            If iNotPlay Then
                wm.Controls.pause
            Else
                wm.Controls.play
            End If
            iNotPlay = Not iNotPlay
            KeyCode = 0
    End Select
End Sub

Private Sub wm_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    If iNotPlay Then
        wm.Controls.pause
    Else
        wm.Controls.play
    End If
    iNotPlay = Not iNotPlay
End Sub

