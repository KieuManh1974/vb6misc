VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WMPLibCtl.WindowsMediaPlayer wm 
      Height          =   3675
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3600
      URL             =   "C:\Documents and Settings\All Users.WINDOWS\Documents\My Videos\My Deliveries\iplayer_live\BBC_1_Magic_Forest_16x9.wmv"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6350
      _cy             =   6482
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IKeyboardHook

' C:\Documents and Settings\All Users.WINDOWS\Documents\My Videos\My Deliveries\iplayer_live\BBC_1_Magic_Forest_16x9.wmv
Private iNotPlay As Boolean

Private Sub Form_Load()
    Set KeyboardHandler.KeyboardHook = Me
    HookKeyboard
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookKeyboard
End Sub


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
    
    IKeyboardHook_ProcessKey iLastPosition + 48, 0
       
    'lTVTool = OpenApplication("D:\Program Files\TVTool 8\TVTOOL.exe", "D:\Program Files\TVTool 8\")
    'Wait 5, lTVTool
    'CloseApplication lTVTool
End Sub



Private Function IKeyboardHook_ProcessKey(ByVal lKeyCode As Long, ByVal lShift As Long) As Boolean
    Dim iIndex As Integer
    Dim lStepSize As Long
    Static lType As Long
    
    On Error Resume Next
    
    Select Case lKeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
            IKeyboardHook_ProcessKey = True
    End Select
    
    If (lShift And 2) = 0 Then
        lStepSize = 1
    Else
        lStepSize = 32
    End If
        
    Select Case lType
        Case 0
            Select Case lShift And 1
                Case 0
                    Select Case lKeyCode
                        Case vbKeyLeft
                            wm.Left = wm.Left - lStepSize
                        Case vbKeyRight
                            wm.Left = wm.Left + lStepSize
                        Case vbKeyUp
                            wm.Top = wm.Top - lStepSize
                        Case vbKeyDown
                            wm.Top = wm.Top + lStepSize
                    End Select
                Case 1
                    Select Case lKeyCode
                        Case vbKeyLeft
                            Me.Left = Me.Left - lStepSize * Screen.TwipsPerPixelX
                        Case vbKeyRight
                            Me.Left = Me.Left + lStepSize * Screen.TwipsPerPixelX
                        Case vbKeyUp
                            Me.Top = Me.Top - lStepSize * Screen.TwipsPerPixelY
                        Case vbKeyDown
                            Me.Top = Me.Top + lStepSize * Screen.TwipsPerPixelY
                    End Select
            End Select
        Case 1
            Select Case lShift And 1
                Case 0
                    Select Case lKeyCode
                        Case vbKeyLeft
                            wm.Left = wm.Left - lStepSize
                            wm.Width = wm.Width + 2 * lStepSize
                        Case vbKeyRight
                            wm.Left = wm.Left + lStepSize
                            wm.Width = wm.Width - 2 * lStepSize
                        Case vbKeyUp
                            wm.Top = wm.Top - lStepSize
                            wm.Height = wm.Height + 2 * lStepSize
                        Case vbKeyDown
                            wm.Top = wm.Top + lStepSize
                            wm.Height = wm.Height - 2 * lStepSize
                    End Select
                Case 1
                    Select Case lKeyCode
                        Case vbKeyLeft
                            Me.Left = Me.Left - lStepSize * Screen.TwipsPerPixelX
                            Me.Width = Me.Width + 2 * lStepSize * Screen.TwipsPerPixelX
                        Case vbKeyRight
                            Me.Left = Me.Left + lStepSize * Screen.TwipsPerPixelX
                            Me.Width = Me.Width - 2 * lStepSize * Screen.TwipsPerPixelX
                        Case vbKeyUp
                            Me.Top = Me.Top - lStepSize * Screen.TwipsPerPixelY
                            Me.Height = Me.Height + 2 * lStepSize * Screen.TwipsPerPixelY
                        Case vbKeyDown
                            Me.Top = Me.Top + lStepSize * Screen.TwipsPerPixelY
                            Me.Height = Me.Height - 2 * lStepSize * Screen.TwipsPerPixelY
                    End Select
            End Select
    End Select
        

    Select Case lKeyCode
        Case vbKeyZ
            lType = 0
        Case vbKeyX
            lType = 1
        Case vbKeyP, vbKeyQ, vbKeyW, vbKeyE, vbKeyR, vbKeyT, vbKeyY, vbKeyU, vbKeyI, vbKeyO
            iIndex = InStr("PQWERTYUIO", Chr$(lKeyCode)) - 1

            SaveSetting "VideoPlayer", "Screen" & iIndex, "w", Me.Width
            SaveSetting "VideoPlayer", "Screen" & iIndex, "h", Me.Height
            SaveSetting "VideoPlayer", "Screen" & iIndex, "y", Me.Top
            SaveSetting "VideoPlayer", "Screen" & iIndex, "x", Me.Left

            SaveSetting "VideoPlayer", "Position" & iIndex, "x", wm.Left
            SaveSetting "VideoPlayer", "Position" & iIndex, "y", wm.Top
            SaveSetting "VideoPlayer", "Position" & iIndex, "w", wm.Width
            SaveSetting "VideoPlayer", "Position" & iIndex, "h", wm.Height
            SaveSetting "VideoPlayer", "LastPosition", "p", iIndex
        Case 48 To 57
            Me.Left = GetSetting("VideoPlayer", "Screen" & lKeyCode - 48, "x", Me.Left)
            Me.Top = GetSetting("VideoPlayer", "Screen" & lKeyCode - 48, "y", Me.Top)
            Me.Width = GetSetting("VideoPlayer", "Screen" & lKeyCode - 48, "w", Me.Width)
            Me.Height = GetSetting("VideoPlayer", "Screen" & lKeyCode - 48, "h", Me.Height)

            wm.Left = GetSetting("VideoPlayer", "Position" & lKeyCode - 48, "x", wm.Left)
            wm.Top = GetSetting("VideoPlayer", "Position" & lKeyCode - 48, "y", wm.Top)
            wm.Width = GetSetting("VideoPlayer", "Position" & lKeyCode - 48, "w", wm.Width)
            wm.Height = GetSetting("VideoPlayer", "Position" & lKeyCode - 48, "h", wm.Height)

            SaveSetting "VideoPlayer", "LastPosition", "p", lKeyCode - 48
        Case 27
            Unload Me
        Case vbKeyM
            Me.WindowState = vbMaximized
        Case vbKeyN
            Me.WindowState = vbNormal
        Case vbKeySpace
            If iNotPlay Then
                wm.Controls.pause
            Else
                wm.Controls.play
            End If
            iNotPlay = Not iNotPlay
    End Select
End Function


