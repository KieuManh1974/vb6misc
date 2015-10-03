VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDXSound 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPlay 
      Interval        =   100
      Left            =   5460
      Top             =   7560
   End
   Begin VB.PictureBox picWave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   180
      ScaleHeight     =   465
      ScaleWidth      =   5385
      TabIndex        =   26
      Top             =   1260
      Width           =   5415
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.PictureBox picWave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   1
      Left            =   180
      ScaleHeight     =   465
      ScaleWidth      =   5385
      TabIndex        =   25
      Top             =   720
      Width           =   5415
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.PictureBox picPos 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   60
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   22
      Top             =   5400
      Width           =   5715
   End
   Begin VB.PictureBox picWave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   180
      ScaleHeight     =   465
      ScaleWidth      =   5385
      TabIndex        =   21
      Top             =   180
      Width           =   5415
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   3420
      TabIndex        =   12
      Top             =   3660
      Width           =   2355
      Begin VB.CommandButton Command1 
         Caption         =   "CLOSE SAMPLE"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   660
         Width           =   2115
      End
      Begin VB.CheckBox ckLoop 
         Caption         =   "LOOP"
         Height          =   315
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdWave 
         Caption         =   "STOP"
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdWave 
         Caption         =   "PLAY"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   3540
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Third Wave File"
      Height          =   315
      Index           =   2
      Left            =   3420
      TabIndex        =   11
      Top             =   2760
      Width           =   2355
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Second Wave File"
      Height          =   315
      Index           =   1
      Left            =   3420
      TabIndex        =   10
      Top             =   2400
      Width           =   2355
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select First Wave File"
      Height          =   315
      Index           =   0
      Left            =   3420
      TabIndex        =   9
      Top             =   2040
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   2895
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   1515
      Begin MSComctlLib.Slider scrDir 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   840
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin MSComctlLib.Slider scrDir 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   19
         Top             =   1620
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin MSComctlLib.Slider scrDir 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   2460
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin VB.Label lblPan 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (1)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label lblPan 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (2)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label lblPan 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (3)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   2100
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume"
      Height          =   2895
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   1515
      Begin MSComctlLib.Slider scrVol 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin MSComctlLib.Slider scrVol 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   1620
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin MSComctlLib.Slider scrVol 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   17
         Top             =   2460
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
         TickFrequency   =   100
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (3)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (2)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave (1)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   1155
      End
   End
   Begin VB.Label lbHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use Mouse Left Button To Move Sound 1, Use Right Mouse Button To Move Sound 2 And Use Shift+Right Mouse utton To Move Sound 3"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   60
      TabIndex        =   24
      Top             =   4860
      Width           =   5715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0FFFF&
      Height          =   1695
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   5715
   End
End
Attribute VB_Name = "frmDXSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements DirectXEvent
Dim bufferDesc(2) As DSBUFFERDESC 'a new object that when filled in is passed to the DS object to describe
Dim WAVEFORMAT(2) As WAVEFORMATEX
Dim m_pos(2) As D3DVECTOR
Dim m_bMouseDown As Boolean
Dim intPlayMode As Integer
Dim dxEventID(2) As Long
'this array notify the position when playng a wave file
Dim dxEVNT(1) As DSBPOSITIONNOTIFY

Private Sub DrawWaves(ByVal Ind As Integer)
   'Graph the waveform
   Dim x As Long               ' current X position
   Dim leftYOffset As Long     ' Y offset for left channel graph
   Dim rightYOffset As Long    ' Y offset for right channel graph
   Dim curLeftY As Long        ' current left channel Y value
   Dim curRightY As Long       ' current right channel Y value
   Dim lastX As Long           ' last X position
   Dim lastLeftY As Long       ' last left channel Y value
   Dim lastRightY As Long      ' last right channel Y value
   Dim maxAmplitude As Long    ' the maximum amplitude for a wavegraph on the form
   Dim leftVol As Double       ' buffer for retrieving the left volupicwave(ind) level
   Dim rightVol As Double      ' buffer for retrieving the right volupicwave(ind) level
   Dim ScaleFactor As Double   ' samples per pixel on the wave graph
   Dim xStep As Double         ' pixels per sample on the wave graph
   Dim curSample As Long       ' current sample number
   
   ' clear the screen
   picWave(Ind).Cls
   
   ' if no file is loaded, don't try to draw graph
   If Not GetFileInformation(lblFile(Ind).Tag) Then
       Exit Sub
   End If
   
   ' calculate drawing parameters
   ScaleFactor = (mdlWave.numSamples - 0) / picWave(Ind).Width
   If (ScaleFactor < 1) Then
       xStep = 1 / ScaleFactor
   Else
       xStep = 1
   End If
   
   ' Draw the graph
   If (mdlWave.format.nChannels = 2) Then
      'Draw Left and Right Channel (Stereo)
      maxAmplitude = picWave(Ind).Height / 4
      leftYOffset = maxAmplitude
      rightYOffset = maxAmplitude * 3
       
      For x = 0 To picWave(Ind).Width Step xStep
         curSample = ScaleFactor * x + mdlWave.drawFrom
         If (mdlWave.format.wBitsPerSample = 16) Then
             GetStereo16Sample curSample, leftVol, rightVol
         Else
             GetStereo8Sample curSample, leftVol, rightVol
         End If
         curRightY = CLng(rightVol * maxAmplitude)
         curLeftY = CLng(leftVol * maxAmplitude)
         picWave(Ind).Line (lastX, leftYOffset + lastLeftY)-(x, curLeftY + leftYOffset)
         picWave(Ind).Line (lastX, rightYOffset + lastRightY)-(x, curRightY + rightYOffset)
         lastLeftY = curLeftY
         lastRightY = curRightY
         lastX = x
      Next
   Else
      ''Draw Single Channel (mono)
      maxAmplitude = picWave(Ind).Height / 2
      leftYOffset = maxAmplitude
      
      For x = 0 To picWave(Ind).Width Step xStep
         curSample = ScaleFactor * x + mdlWave.drawFrom
         If (mdlWave.format.wBitsPerSample = 16) Then
             GetMono16Sample curSample, leftVol
         Else
             GetMono8Sample curSample, leftVol
         End If
         curLeftY = CLng(leftVol * maxAmplitude)
         picWave(Ind).Line (lastX, leftYOffset + lastLeftY)-(x, curLeftY + leftYOffset)
         lastLeftY = curLeftY
         lastX = x
      Next
   End If

End Sub


Private Sub DrawTriangle(col As Long, x As Integer, z As Integer, ByVal a As Single)
    
    Dim x1 As Integer
    Dim z1 As Integer
    Dim x2 As Integer
    Dim z2 As Integer
    Dim x3 As Integer
    Dim z3 As Integer
    
    a = 3.141 * (a - 90) / 180
    Dim q As Integer
    q = 10
    
    x1 = q * Sin(a) + x
    z1 = q * Cos(a) + z
    
    x2 = q * Sin(a + 3.141 / 1.3) + x
    z2 = q * Cos(a + 3.141 / 1.3) + z
    
    x3 = q * Sin(a - 3.141 / 1.3) + x
    z3 = q * Cos(a - 3.141 / 1.3) + z
    
    
    
    picPos.Line (x1, z1)-(x2, z2), col
    picPos.Line (x1, z1)-(x3, z3), col
    picPos.Line (x2, z2)-(x3, z3), col
End Sub


Private Sub LoadWave(ByVal Ind As Integer)

    'These settings should do for almost any app....
    bufferDesc(Ind).lFlags = DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_STATIC Or DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    WAVEFORMAT(Ind).nSize = LenB(WAVEFORMAT(Ind))
    WAVEFORMAT(Ind).nFormatTag = WAVE_FORMAT_PCM
    WAVEFORMAT(Ind).nChannels = 2
    WAVEFORMAT(Ind).lSamplesPerSec = 22050
    WAVEFORMAT(Ind).nBitsPerSample = 16
    WAVEFORMAT(Ind).nBlockAlign = WAVEFORMAT(Ind).nBitsPerSample / 8 * WAVEFORMAT(Ind).nChannels
    WAVEFORMAT(Ind).lAvgBytesPerSec = WAVEFORMAT(Ind).lSamplesPerSec * WAVEFORMAT(Ind).nBlockAlign
    
    
    Dim Cn As Integer
    Dim sFile As String
    
    sFile = lblFile(Ind).Tag
    If sFile <> "" Then
        Set m_dsBuffer(Ind) = m_ds.CreateSoundBufferFromFile(sFile, bufferDesc(Ind), WAVEFORMAT(Ind))
    End If
    
    'checks for any errors
    If err.Number <> 0 Then 'basically, generate an error message if the number is anything but 0(0=no error)
        MsgBox "unable to find " + sFile
        End
    End If
        
    'create Direct3D sound buffer for positional audio
    Set m_ds3DBuffer(Ind) = m_dsBuffer(Ind).GetDirectSound3DBuffer
    
    m_ds3DBuffer(Ind).SetConeAngles DS3D_MINCONEANGLE, 100, DS3D_IMMEDIATE
    m_ds3DBuffer(Ind).SetConeOutsideVolume -400, DS3D_IMMEDIATE
    

    'position our sound
    m_ds3DBuffer(Ind).SetPosition m_pos(i).x / 50, 0, m_pos(i).z / 50, DS3D_IMMEDIATE
    
    scrDir_Change Ind
    scrVol_Change Ind
    'i destroy every previus events
    If dxEventID(Ind) <> 0 Then
      m_dx.DestroyEvent dxEventID(Ind)
    End If
    
    'I create a new event that the end of
    'the execution of the sound notifies me
    dxEventID(Ind) = m_dx.CreateEvent(Me)
    dxEVNT(0).hEventNotify = dxEventID(Ind)
    dxEVNT(0).lOffset = bufferDesc(Ind).lBufferBytes - 1
        
    m_dsBuffer(Ind).SetNotificationPositions 1, dxEVNT()
End Sub

Private Sub ckLoop_Click()
    intPlayMode = ckLoop.Value
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    cmDlg.Filter = "File Wave (*.Wav)|*.Wav"
    cmDlg.ShowOpen
    
    If cmDlg.FileName = "" Then Exit Sub
    Select Case Index
    Case 0
        'FIRST FILE SELECTED
        lblFile(0).BackColor = vbYellow
        lblFile(0).Tag = cmDlg.FileName
    Case 1
        'SECOND FILE SELECTED
        lblFile(1).BackColor = vbYellow
        lblFile(1).Tag = cmDlg.FileName
    Case 2
        'THIRD FILE SELECTED
        lblFile(2).BackColor = vbYellow
        lblFile(2).Tag = cmDlg.FileName
    End Select
    DrawWaves Index
    LoadWave Index
End Sub

Private Sub cmdWave_Click(Index As Integer)
    Dim Cn As Integer
    Select Case Index
    Case 0  'PLAY
        
        For Cn = 0 To 2
            If m_dsBuffer(Cn) Is Nothing = False Then
                m_dsBuffer(Cn).Play intPlayMode
            End If
        Next
    Case 1  'STOP
        For Cn = 0 To 2
            If m_dsBuffer(Cn) Is Nothing = False Then
               m_dsBuffer(Cn).Stop
               m_dsBuffer(Cn).SetCurrentPosition 0
               Line1(Cn).x1 = 0
               Line1(Cn).x2 = 0
            End If
        Next
    Case 2  'record
        
    End Select
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
    Select Case eventid
    Case dxEventID(0)
        Line1(0).x1 = 0: Line1(0).x2 = 0
    Case dxEventID(1)
        Line1(1).x1 = 0: Line1(1).x2 = 0
    Case dxEventID(2)
        Line1(2).x1 = 0: Line1(2).x2 = 0
    End Select
End Sub

Private Sub Form_Load()
    Dim Cn As Integer
    
    For Cn = 0 To 2
        scrVol(Cn).Max = 0
        scrVol(Cn).Min = -5000
        scrVol(Cn).LargeChange = 20
        scrVol(Cn).SmallChange = 255
     
        scrDir(Cn).Max = 360
        scrDir(Cn).Min = -360
        scrDir(Cn).LargeChange = 30
        scrDir(Cn).SmallChange = 5
        
        scrVol(Cn).Value = -1000
        scrDir(Cn).Value = -90
     Next
    intPlayMode = 0
    Me.Show
    On Local Error Resume Next
    'First we have to create a DSound object, this must be done before any features can be used.
    'It must also be done before we set the cooperativelevel or create any buffers.
    Set m_ds = m_dx.DirectSoundCreate("")
    'This checks for any errors, if there are no errors the user has got DX7 and a functional sound card
    If err.Number <> 0 Then
        MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
        End
    End If
    'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS
    'associating our DS object with our window is important. This tells windows to stop
    'other sounds from interfering with ours, and ours not to interfere with other apps.
    'The sounds will only be played when the from has got focus....
    'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card
    'Needed for games
    'DSSCL_NORMAL=cooperates with other apps, shares resources
    'Good for general windows multimedia apps.
    'DSSCL_PRIORITY seems to be the better setting out for audio applications
    m_ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    
    '****
    'Dim primDesc As DSBUFFERDESC, format As WAVEFORMATEX
    
    'primDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
    'Set m_dsPrimaryBuffer = m_ds.CreateSoundBuffer(primDesc, format)
    'Set m_dsListener = m_dsPrimaryBuffer.GetDirectSound3DListener()
    '*****
    
    m_pos(0).x = 0:  m_pos(0).z = 50
    m_pos(1).x = 10:  m_pos(1).z = 50
    m_pos(2).x = -10:  m_pos(2).z = 50
    
    DrawElements


End Sub
Sub ReDrawPosition(i As Integer, x As Single, z As Single)
    m_pos(i).x = x - picPos.ScaleWidth / 2
    m_pos(i).z = z - picPos.ScaleHeight / 2
    
    DrawElements
    
    'the zero at the end indicates we want the postion updated immediately
    If m_ds3DBuffer(i) Is Nothing Then Exit Sub
    
    m_ds3DBuffer(i).SetPosition m_pos(i).x / 50, 0, m_pos(i).z / 50, DS3D_IMMEDIATE
    
    

End Sub

Private Sub SetPosition(ByVal Ind As Integer, ByVal xPos As Double)
    Dim ScaleFactor As Double
    Dim xStep As Double
    Dim nMove As Double
    
    ScaleFactor = bufferDesc(Ind).lBufferBytes / picWave(Ind).Width
    xStep = ((CurPos.lPlay / picWave(Ind).Width) * picWave(Ind).Width) / ScaleFactor
    
    Line1(Ind).x1 = xStep
    Line1(Ind).x2 = Line1(Ind).x1
End Sub
Private Sub DrawElements()
    Dim x As Integer
    Dim z As Integer
    
    picPos.Cls
    
    'listener is in center and is white
    DrawTriangle vbWhite, picPos.ScaleWidth / 2, picPos.ScaleHeight / 2, 90
    
    'draw sound 1 as yellow
    x = CInt(m_pos(0).x) + picPos.ScaleWidth / 2
    z = CInt(m_pos(0).z) + picPos.ScaleHeight / 2
    DrawTriangle vbYellow, x, z, scrDir(0).Value
    
    'draw sound2 as Green
    x = CInt(m_pos(1).x) + picPos.ScaleWidth / 2
    z = CInt(m_pos(1).z) + picPos.ScaleHeight / 2
    DrawTriangle vbGreen, x, z, scrDir(1).Value
    
    'draw sound 3 as cyan
    x = CInt(m_pos(2).x) + picPos.ScaleWidth / 2
    z = CInt(m_pos(2).z) + picPos.ScaleHeight / 2
    DrawTriangle vbCyan, x, z, scrDir(2).Value
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    
    For i = 0 To UBound(dxEventID)
        DoEvents
        If dxEventID(i) Then m_dx.DestroyEvent dxEventID(i)
    Next
    On Error Resume Next
    Close #f
End Sub

Private Sub picPos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer
    i = 0
    If Button = 2 And Shift = 0 Then
        i = 1
    ElseIf Button = 2 And Shift = 1 Then
        i = 2
    End If
    ReDrawPosition i, x, Y
    m_bMouseDown = True

End Sub

Private Sub picPos_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer
    If m_bMouseDown = False Then Exit Sub
    i = 0
    If Button = 2 And Shift = 0 Then
        i = 1
    ElseIf Button = 2 And Shift = 1 Then
        i = 2
    End If
    ReDrawPosition i, x, Y
End Sub

Private Sub picPos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    m_bMouseDown = False
End Sub

Private Sub picPos_Paint()
    DrawElements
End Sub


Private Sub scrDir_Change(Index As Integer)
    'fist we must calculate a vector of what direction
    'the sound is traveling in.
    '
    Dim x As Single
    Dim z As Single
    'we take the current angle in degrees convert to radians
    'and get the cos or sin to find the direction from an angle
    x = 5 * Cos(3.141 * scrDir(0).Value / 180)
    z = 5 * Sin(3.141 * scrDir(0).Value / 180)
    
    'Update the UI
    DrawElements
    
    If m_dsBuffer(Index) Is Nothing Then Exit Sub
    
    'the zero at the end indicates we want the postion updated immediately
    m_ds3DBuffer(Index).SetConeOrientation x, 0, z, DS3D_IMMEDIATE


End Sub

Private Sub scrVol_Change(Index As Integer)
    If m_dsBuffer(Index) Is Nothing Then Exit Sub
    m_dsBuffer(Index).SetVolume scrVol(Index).Value
End Sub

Private Sub scrVol_Scroll(Index As Integer)
    scrVol_Change (Index)
End Sub

Private Sub tmrPlay_Timer()
    Dim Cn As Integer
    Dim Cpos As Double
    For Cn = 0 To 2
        If m_dsBuffer(Cn) Is Nothing = False Then
            If m_dsBuffer(Cn).GetStatus = DSBSTATUS_PLAYING Then
                m_dsBuffer(Cn).GetCurrentPosition CurPos
                ReDim wBuffer(CurPos.lPlay - 1)
                m_dsBuffer(Cn).ReadBuffer 0, CurPos.lPlay, wBuffer(0), DSBLOCK_DEFAULT
                Cpos = CurPos.lPlay
                SetPosition Cn, Cpos
            End If
        End If
    Next
End Sub
