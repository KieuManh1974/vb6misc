VERSION 5.00
Begin VB.Form frmMidi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Midi IN, OUT & THRU"
   ClientHeight    =   1635
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPorts 
      Caption         =   "Midi keyboard IN -> callback -> OUT"
      Height          =   240
      Index           =   3
      Left            =   3510
      TabIndex        =   6
      Top             =   900
      Width           =   3885
   End
   Begin VB.OptionButton optPorts 
      Caption         =   "Midi keyboard IN <- connect -> OUT = THRU"
      Height          =   240
      Index           =   2
      Left            =   3510
      TabIndex        =   5
      Top             =   630
      Width           =   3885
   End
   Begin VB.OptionButton optPorts 
      Caption         =   "Midi OUT only with mouse"
      Height          =   240
      Index           =   1
      Left            =   3510
      TabIndex        =   4
      Top             =   360
      Width           =   2355
   End
   Begin VB.OptionButton optPorts 
      Caption         =   "All ports close"
      Height          =   240
      Index           =   0
      Left            =   3510
      TabIndex        =   3
      Top             =   90
      Value           =   -1  'True
      Width           =   2805
   End
   Begin VB.PictureBox picKlav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   45
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   1260
      Width           =   7875
   End
   Begin VB.CommandButton cmdAllNotesOff 
      Caption         =   "All Notes Off"
      Height          =   270
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "All Notes Off"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.Label lblMseNote 
      Alignment       =   2  'Center
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5715
      TabIndex        =   7
      Top             =   375
      Width           =   825
   End
   Begin VB.Label lblMidiInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Midi INFO"
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   45
      Width           =   2715
   End
End
Attribute VB_Name = "frmMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is the main startup form
Option Explicit
Dim CurKeyID As Long          ' remember note on for note off
Dim Send As Boolean
' generate piano klavier in picture
Public Sub MakePiano(pic As PictureBox)
   Dim wX1 As Long, wY1 As Long
   Dim wdX As Long, wdY As Long
   Dim zX1 As Long, zY1 As Long
   Dim zdX As Long, zdY As Long
   Dim AaWTs As Long                   ' count white keys
   Dim I As Long                       ' counter
   
   wX1 = 0: wY1 = 0: wdX = 7: wdY = 22 ' witte toets
   zX1 = 5: zY1 = 0: zdX = 4: zdY = 16 ' zwarte

   AaWTs = (128 / 12) * 7

   pic.Width = AaWTs * wdX * 15
   pic.AutoRedraw = True
   
   ' make 1st white key & copy other white keys
   pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(15), BF
   pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(0), B
   For I = 0 To AaWTs - 1
      BitBlt pic.hDC, wX1 + I * wdX, wY1, wdX, wdY + 1, pic.hDC, wX1, wY1, SRCCOPY
   Next I
      
   ' 1st black & copy other
   pic.Line (zX1, zY1)-Step(zdX, zdY), QBColor(0), BF
   For I = 1 To AaWTs - 1
      If Mid("110111", (I Mod 7) + 1, 1) = "1" Then
         BitBlt pic.hDC, zX1 + I * wdX, zY1, zdX + 1, zdY, pic.hDC, zX1, zY1, SRCCOPY
         End If
   Next I
   
   pic.Line (pic.ScaleWidth - 1, wY1)-Step(0, wdY), QBColor(0)
   pic.Picture = pic.Image
   pic.AutoRedraw = False
End Sub


Public Sub ShowNote(ByVal Nr As Long, OnOff As Long)
   Dim octave As Long, note As Long, bw As Long
   Dim X As Long, Y As Long, s As Long
   Dim color As Long
   
   octave = (Nr \ 12)
   note = Nr Mod 12
   bw = Choose(note + 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0) ' black or white
   X = octave * 49 + Choose(note + 1, 0, 3, 7, 10, 14, 21, 24, 28, 31, 35, 38, 42, 49)
   If bw = 1 Then
      Y = 11: X = X + 3: s = 2 ' black key
      color = IIf(OnOff = 1, QBColor(15), 0)
      Else
      Y = 17: X = X + 2: s = 3 ' white key
      color = IIf(OnOff = 1, 0, QBColor(15))
      End If
   picKlav.ForeColor = color
   picKlav.FillColor = color
   picKlav.Line (X, Y)-Step(s, s), color, BF
End Sub

Private Sub cmdAllNotesOff_Click()
   If Send = False Then Exit Sub
   midiMessageOut = CONTROLLER_CHANGE + CurChannel
   midiData1 = &H7B
   midiData2 = CByte(0)
   SendMidiShortOut
End Sub

Private Sub Form_Load()
   App.Title = "Midi IN, OUT & THRU"
   mMPU401OUT = 256 ' =empty
   mMPU401IN = 256
   CurChannel = 0
   MakePiano picKlav
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MidiIN_Port "close"
   MidiOUT_Port "close"
   End
End Sub

Private Sub optPorts_Click(Index As Integer)
   MidiIN_Port "close"
   MidiOUT_Port "close"
   Select Case Index
        Case 0
           lblMidiInfo.Caption = "All ports closed"
        Case 1
           MidiOUT_Port "open"
           lblMidiInfo.Caption = "Mouse input only"
        Case 2
           MidiTHRU_Port "open"
           lblMidiInfo.Caption = "Midi THRU"
        Case 3
           MidiOUT_Port "open"
           MidiIN_Port "open"
   End Select
   Send = IIf(Index = 0, False, True)
End Sub

Private Sub picKlav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim oct As Long
   Dim No As Long
   Dim mX As Single
   
   oct = X \ 49
   If picKlav.Point(X, Y) = 0 And Y < 17 Then
      mX = X - 4
      No = oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
      Else
      mX = X
      No = oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
      End If
   
   CurKeyID = No
   ShowNote No, 1
   lblMseNote.Caption = isNote(No)
   If Send = True Then
      midiMessageOut = NOTE_ON + CurChannel
      midiData1 = No
      midiData2 = 120
      SendMidiShortOut
      End If
End Sub

Private Sub picKlav_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Send = True Then
      midiMessageOut = NOTE_OFF + CurChannel
      midiData1 = CurKeyID
      midiData2 = 0
      SendMidiShortOut
      End If
   ShowNote CurKeyID, 0
End Sub


