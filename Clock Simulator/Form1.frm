VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function playa Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = 0
Const SND_ASYNC = 1
Const SND_NODEFAULT = 2
Const SND_MEMORY = 4
Const SND_LOOP = 8
Const SND_NOSTOP = 16

Private mlTick As Long

Private sTick As String
Private sSingleChime As String
Private sStartMultiChime As String
Private sEndMultiChime As String
Private sMidChime As String

Private Sub Form_Load()
    sTick = App.Path & "\tick2.wav"
    sSingleChime = App.Path & "\singlechime.wav"
    sStartMultiChime = App.Path & "\multichimestart.wav"
    sEndMultiChime = App.Path & "\multichimeend.wav"
    sMidChime = App.Path & "\midchime.wav"
    
    Dim dPreviousTime As Date
    dPreviousTime = Time
    While Time = dPreviousTime
        dPreviousTime = Time
    Wend
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()

    Dim lSecond As Long
    Dim lMinute As Long
    Dim lHour As Long
    Dim dTime As Date
    Dim x As Long
    
    dTime = Time
    
    lSecond = Format$(dTime, "SS")
    lMinute = Format$(dTime, "N")
    lHour = Format$(dTime, "HH") Mod 12
    If lHour = 0 Then
        lHour = 12
    End If
    
    mlTick = lSecond Mod 4
    'Debug.Print dTime
        
'    If lSecond = 0 Then
'        If lMinute = 0 Then
'            If lHour = 1 Then
'                playa sSingleChime, 0
'            Else
'                playa sStartMultiChime, 0
'                For x = 1 To lHour - 2
'                    playa sMidChime, 0
'                Next
'                playa sEndMultiChime, 0
'            End If
'        ElseIf lMinute = 30 Then
'            playa sStartMultiChime, 0
'            playa sEndMultiChime, 0
'        ElseIf (lMinute Mod 15) = 0 Then
'            playa sSingleChime, 0
'        Else
'            If mlTick = 0 Then
'                playa sTick, 1
'            End If
'        End If
'    Else
        If mlTick = 0 Then
            playa sTick, 1
        End If
'    End If
    
    
End Sub
