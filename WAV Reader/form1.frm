VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Guitar"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13950
   ForeColor       =   &H0002B7C6&
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Display 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   14295
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   949
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   901
      TabIndex        =   12
      Top             =   480
      Width           =   13575
   End
   Begin VB.PictureBox FretBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   805
      TabIndex        =   11
      Top             =   10920
      Visible         =   0   'False
      Width           =   12135
   End
   Begin VB.TextBox txtFile 
      Height          =   288
      Left            =   6960
      TabIndex        =   9
      Text            =   "c:\main\projects\sampler\track1.wav"
      Top             =   120
      Width           =   3972
   End
   Begin VB.CommandButton butStop 
      Caption         =   "¾"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   7.5
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   252
   End
   Begin VB.CommandButton butPlay 
      Caption         =   "„"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   7.5
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   252
   End
   Begin VB.TextBox txtStep 
      Height          =   288
      Left            =   4080
      TabIndex        =   5
      Text            =   "0.01"
      Top             =   120
      Width           =   492
   End
   Begin VB.TextBox bottomA 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   2280
      TabIndex        =   2
      Text            =   "55"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtTime 
      Height          =   288
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "Step"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "File"
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Hz"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "A ="
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DFTSize = 65536

Private Const pi = 3.14159265358979

'Private Const NumberOfNotes = 88 - 1
Private Const NumberOfNotes = 700
Private Const NoteDivision = 96
Private Const NoteWidth = 1
Private DataStart As Long
Private SampleRate As Long

Private ScaleColour(11) As Long


Private KeyPowerReal(NumberOfNotes) As Double
Private KeyPowerImag(NumberOfNotes) As Double

'Private LastKeyPhase(NumberOfNotes) As Double

Private Const Spacing = 15
Private Const LeftMargin = 40
Private Const RightMargin = 950
Private Const TopMargin = 10

Private Playing As Boolean

Private Sub butPlay_Click()
    Dim lCounter As Long
    
    Playing = True
    
    Display.Cls
    ContinuousRead3 (txtTime.Text) * SampleRate
End Sub

Private Sub butStop_Click()
    Playing = False
End Sub


Private Sub LoadFile()
    Dim chunkdescriptor As String * 4
    Dim chunksize As Long
    Dim wave As String * 4
    Dim fmt As String * 4
    Dim subchunk1size As Long
    Dim audioformat As Integer
    Dim numchanels As Integer
    Dim byterate As Long
    Dim blockalign As Integer
    Dim bitspersample As Integer
    Dim data As String * 4
    Dim subchunk2size As Long
    
    If Dir(txtFile.Text) <> "" Then
        txtFile.ForeColor = vbBlack
        
        ' Open the leading file
        Close #1
        Open txtFile.Text For Binary As #1
        Get #1, , chunkdescriptor
        Get #1, , chunksize
        Get #1, , wave
        Get #1, , fmt
        Get #1, , subchunk1size
        Get #1, , audioformat
        Get #1, , numchanels
        Get #1, , SampleRate
        Get #1, , byterate
        Get #1, , blockalign
        Get #1, , bitspersample
        Get #1, , data
        Get #1, , subchunk2size
        DataStart = Seek(1)
        
        ' Open the following file
        Close #2
        Open txtFile.Text For Binary As #2
        Seek #2, DataStart
    End If
    
End Sub


Private Sub ContinuousRead(SamplePos As Long)
'    Dim iSample As Long
'    Dim iSignal As Integer
'    Dim iKey As Integer
'    Dim iShowKeyboard As Integer
'    Dim iOutput As Integer
'    Dim iDirection As Integer
'    Dim iFileNum As Integer
'    Dim lSample(1 To 2) As Long
'
'    Dim ResetValue(NumberOfNotes) As Double
'    Dim CounterReal(NumberOfNotes, 1 To 2) As Double
'    Dim CounterImag(NumberOfNotes, 1 To 2) As Double
'    Dim ValueReal(NumberOfNotes, 1 To 2) As Integer
'    Dim ValueImag(NumberOfNotes, 1 To 2) As Integer
'    Dim SumPowerReal(NumberOfNotes) As Long
'    Dim SumPowerImag(NumberOfNotes) As Long
'    Dim dFrequency As Double
'
'    'Load Reset values
'    For iKey = 0 To NumberOfNotes
'        dFrequency = Val(bottomA.Text) * 2 ^ (iKey / 12)
'        ResetValue(iKey) = SampleRate / (dFrequency * 2)
'        CounterImag(iKey, 1) = ResetValue(iKey)
'        CounterReal(iKey, 1) = ResetValue(iKey) / 2
'        ValueImag(iKey, 1) = 1
'        ValueReal(iKey, 1) = 1
'        CounterImag(iKey, 2) = ResetValue(iKey)
'        CounterReal(iKey, 2) = ResetValue(iKey) / 2
'        ValueImag(iKey, 2) = 1
'        ValueReal(iKey, 2) = 1
'        SumPowerReal(iKey) = 0
'        SumPowerImag(iKey) = 0
'        KeyPowerReal(iKey) = 0
'        KeyPowerImag(iKey) = 0
'    Next
'
'    Seek #1, SamplePos * 2 + DataStart
'    Seek #2, SamplePos * 2 + DataStart
'
'    iShowKeyboard = 0
'
'    For lSample(2) = 0 To DFTSize - 1
'        Get #2, , iSignal
'
'        For iKey = 0 To NumberOfNotes
'            SumPowerReal(iKey) = SumPowerReal(iKey) + iSignal * ValueReal(iKey, 2)
'            SumPowerImag(iKey) = SumPowerImag(iKey) + iSignal * ValueImag(iKey, 2)
'
'            CounterImag(iKey, 2) = CounterImag(iKey, 2) - 1
'            If CounterImag(iKey, 2) < 0 Then
'                CounterImag(iKey, 2) = CounterImag(iKey, 2) + ResetValue(iKey)
'                ValueImag(iKey, 2) = -ValueImag(iKey, 2)
'            End If
'
'            CounterReal(iKey, 2) = CounterReal(iKey, 2) - 1
'            If CounterReal(iKey, 2) < 0 Then
'                CounterReal(iKey, 2) = CounterReal(iKey, 2) + ResetValue(iKey)
'                ValueReal(iKey, 2) = -ValueReal(iKey, 2)
'            End If
'        Next
'    Next
'
'    While Playing
'        For iFileNum = 1 To 2
'            iDirection = (iFileNum * 2) - 3
'            Get #iFileNum, , iSignal
'
'            lSample(iFileNum) = lSample(iFileNum) + 1
'
'            For iKey = 0 To NumberOfNotes
'                SumPowerReal(iKey) = SumPowerReal(iKey) + iDirection * iSignal * ValueReal(iKey, iFileNum)
'                SumPowerImag(iKey) = SumPowerImag(iKey) + iDirection * iSignal * ValueImag(iKey, iFileNum)
'
'                CounterImag(iKey, iFileNum) = CounterImag(iKey, iFileNum) - 1
'                If CounterImag(iKey, iFileNum) < 0 Then
'                    CounterImag(iKey, iFileNum) = CounterImag(iKey, iFileNum) + ResetValue(iKey)
'                    ValueImag(iKey, iFileNum) = -ValueImag(iKey, iFileNum)
'                End If
'
'                CounterReal(iKey, iFileNum) = CounterReal(iKey, iFileNum) - 1
'                If CounterReal(iKey, iFileNum) < 0 Then
'                    CounterReal(iKey, iFileNum) = CounterReal(iKey, iFileNum) + ResetValue(iKey)
'                    ValueReal(iKey, iFileNum) = -ValueReal(iKey, iFileNum)
'                End If
'                KeyPowerImag(iKey) = SumPowerImag(iKey)
'                KeyPowerReal(iKey) = SumPowerReal(iKey)
'            Next
'        Next
'
'        iShowKeyboard = iShowKeyboard - 1
'        If iShowKeyboard < 1 Then
'            DisplayFretBoard
'            txtTime = Format((((Seek(1) - DataStart) \ 2) / SampleRate), ".000")
'            DoEvents
'            iShowKeyboard = 100
'        End If
'    Wend
End Sub


Private Sub ContinuousRead2(SamplePos As Long)
    Dim iSample As Long
    Dim iSignal As Integer
    Dim iSignalLeft As Integer
    Dim iSignalRight As Integer
    Dim iKey As Integer
    Dim iShowKeyboard As Long
    Dim iOutput As Integer
    Dim iDirection As Integer
    Dim iFileNum As Integer
    Dim lSample(1 To 2) As Long
    
    Dim CosKey(NumberOfNotes + 20) As Double
    Dim SinKey(NumberOfNotes + 20) As Double
    
    Dim ValueCos(NumberOfNotes + 20, 1 To 2) As Double
    Dim ValueSin(NumberOfNotes + 20, 1 To 2) As Double
    
    Dim dFrequency As Double
    Dim vCos As Double
    Dim vSin As Double
    
    Dim iSamples As Long
    
    'Load Reset values
    For iKey = 0 To NumberOfNotes + 20
        dFrequency = Val(bottomA.Text) * 2 ^ (iKey / 12)
        CosKey(iKey) = Cos((-2 * pi * dFrequency) / SampleRate)
        SinKey(iKey) = Sin((-2 * pi * dFrequency) / SampleRate)
        ValueSin(iKey, 1) = 0
        ValueCos(iKey, 1) = 1
        ValueSin(iKey, 2) = 0
        ValueCos(iKey, 2) = 1
    Next
    For iKey = 0 To NumberOfNotes
        KeyPowerReal(iKey) = 0
        KeyPowerImag(iKey) = 0
    Next
    
    Seek #1, SamplePos * 4 + DataStart
    Seek #2, SamplePos * 4 + DataStart
    
    iShowKeyboard = 0
    
    For lSample(2) = 0 To DFTSize - 1
        Get #2, , iSignalLeft
        Get #2, , iSignalRight
        iSignal = (iSignalLeft + iSignalRight) \ 2
        For iKey = 0 To NumberOfNotes
            KeyPowerReal(iKey) = KeyPowerReal(iKey) + iSignal * (ValueCos(iKey, 2))
            KeyPowerImag(iKey) = KeyPowerImag(iKey) + iSignal * (ValueSin(iKey, 2))
        Next
        
        For iKey = 0 To NumberOfNotes
            vCos = ValueCos(iKey, 2) * CosKey(iKey) - ValueSin(iKey, 2) * SinKey(iKey)
            vSin = ValueSin(iKey, 2) * CosKey(iKey) + ValueCos(iKey, 2) * SinKey(iKey)
            ValueCos(iKey, 2) = vCos
            ValueSin(iKey, 2) = vSin
        Next
    Next
    
    While Playing
        For iFileNum = 1 To 2

            Get #iFileNum, , iSignalLeft
            Get #iFileNum, , iSignalRight
            iSignal = (iSignalLeft + iSignalRight) \ 2
            
            iDirection = (iFileNum * 2) - 3
            iSignal = iSignal * iDirection
            
            'lSample(iFileNum) = lSample(iFileNum) + 1
            
            For iKey = 0 To NumberOfNotes
                KeyPowerReal(iKey) = KeyPowerReal(iKey) + iSignal * (ValueCos(iKey, iFileNum))
                KeyPowerImag(iKey) = KeyPowerImag(iKey) + iSignal * (ValueSin(iKey, iFileNum))
            Next
            
            For iKey = 0 To NumberOfNotes
                vCos = ValueCos(iKey, iFileNum)
                ValueCos(iKey, iFileNum) = ValueCos(iKey, iFileNum) * CosKey(iKey) - ValueSin(iKey, iFileNum) * SinKey(iKey)
                ValueSin(iKey, iFileNum) = ValueSin(iKey, iFileNum) * CosKey(iKey) + vCos * SinKey(iKey)
            Next
        Next
        
        iShowKeyboard = iShowKeyboard - 1
        If iShowKeyboard < 1 Then
            DisplayFretBoard (iSamples)
            txtTime = Format((((Seek(1) - DataStart) \ 4) / SampleRate), ".000")
            DoEvents
            iShowKeyboard = CLng(SampleRate) * Val(txtStep.Text)
            iSamples = iSamples + 1
        End If
    Wend
End Sub






'Private Sub Display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Static oldy As Single
'    Static bnotfirsttime As Boolean
'
'
'    If bnotfirsttime Then
'        Display.DrawMode = vbXorPen
'        Display.Line (0, oldy)-(Display.Width, oldy), vbWhite
'        Display.DrawMode = vbCopyPen
'    Else
'        bnotfirsttime = True
'    End If
'
'    Display.DrawMode = vbXorPen
'    Display.Line (0, Y)-(Display.Width, Y), vbWhite
'    Display.DrawMode = vbCopyPen
'    oldy = Y
'
'End Sub


Private Sub Form_Activate()
    'DrawFretBoard
End Sub

Private Sub Form_Load()
    ScaleColour(0) = vbBlue + vbGreen ' A
    ScaleColour(1) = vbWhite
    ScaleColour(2) = vbGreen               ' B
    ScaleColour(3) = vbRed + vbGreen ' C
    ScaleColour(4) = vbWhite
    ScaleColour(5) = &HFF00FF          ' D
    ScaleColour(6) = vbWhite
    ScaleColour(7) = &H80FF&           ' E
    ScaleColour(8) = &H2B7C6                ' F
    ScaleColour(9) = vbWhite
    ScaleColour(10) = vbRed                ' G
    ScaleColour(11) = vbWhite
    
    LoadFile

End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub txtFile_Change()
    txtFile.ForeColor = vbRed
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Playing = False
        LoadFile
    End If
End Sub

Private Sub DrawFretBoard()
    Dim iString As Integer
    Dim iFret As Integer
    Dim lPos As Long
    
    For iString = 0 To 5
        FretBoard.Line (LeftMargin, iString * Spacing + TopMargin)-(RightMargin, iString * Spacing + TopMargin)
    Next
    
    For iFret = 0 To 18
        lPos = RightMargin - (RightMargin - LeftMargin) / (2 ^ (iFret / 12))
        If iFret = 5 Or iFret = 12 Then
            FretBoard.Line (lPos, TopMargin)-(lPos, TopMargin + 5 * Spacing), RGB(0, 255, 0)
        Else
            FretBoard.Line (lPos, TopMargin)-(lPos, TopMargin + 5 * Spacing)
        End If
    Next
End Sub
'
'Private Sub DrawNote(ByVal iNote As Integer, ByVal lColour As Long)
'    Dim iString As Integer
'    Dim iOffset As Integer
'    Dim iPos As Integer
'    Dim lPos1 As Long
'    Dim lPos2 As Long
'    Dim lPos As Long
'
'    For iString = 0 To 5
'        iOffset = (5 - iString) * 5
'        If iString <= 1 Then
'            iOffset = iOffset - 1
'        End If
'
'        ' If note is displayable for this string, draw it
'        If iNote >= iOffset And iNote <= (iOffset + 18) Then
'            iPos = iNote - iOffset
'            lPos1 = RightMargin - (RightMargin - LeftMargin) / (2 ^ (iPos / 12))
'            lPos2 = RightMargin - (RightMargin - LeftMargin) / (2 ^ ((iPos - 1) / 12))
'            lPos = (lPos1 + lPos2) / 2
'            FretBoard.Line (lPos - 5, iString * Spacing + TopMargin - 5)-(lPos + 5, iString * Spacing + TopMargin + 5), lColour, BF
'            If iPos <> 0 Then
'                FretBoard.Line (lPos - 5, iString * Spacing + TopMargin)-(lPos + 6, iString * Spacing + TopMargin)
'            End If
'        End If
'
'    Next
'
'End Sub



Private Sub ContinuousRead3(SamplePos As Long)
    Dim iSample As Long
    Dim iSignal As Integer
    Dim iSignalLeft As Integer
    Dim iSignalRight As Integer
    Dim iKey As Integer
    Dim iShowKeyboard As Long
    Dim iOutput As Integer
    Dim iDirection As Integer
    Dim iFileNum As Integer
    Dim lSample(1 To 2) As Long
    
    Dim CosKey(NumberOfNotes) As Double
    Dim SinKey(NumberOfNotes) As Double
    
    Dim ValueCos(NumberOfNotes, 1 To 2) As Double
    Dim ValueSin(NumberOfNotes, 1 To 2) As Double
    
    Dim dFrequency As Double
    Dim vCos As Double
    Dim vSin As Double
    
    Dim iSamples As Long
    
    'Load Reset values
    For iKey = 0 To NumberOfNotes
        dFrequency = Val(bottomA.Text) * 2 ^ (iKey / NoteDivision)
        CosKey(iKey) = Cos((-2 * pi * dFrequency) / SampleRate)
        SinKey(iKey) = Sin((-2 * pi * dFrequency) / SampleRate)
        ValueSin(iKey, 1) = 0
        ValueCos(iKey, 1) = 1
        ValueSin(iKey, 2) = 0
        ValueCos(iKey, 2) = 1
    Next
    For iKey = 0 To NumberOfNotes
        KeyPowerReal(iKey) = 0
        KeyPowerImag(iKey) = 0
    Next
    
    Seek #1, SamplePos * 4 + DataStart
    Seek #2, SamplePos * 4 + DataStart
    
    iShowKeyboard = 0
    
    For lSample(2) = 0 To DFTSize - 1
        Get #2, , iSignalLeft
        Get #2, , iSignalRight
        iSignal = (iSignalLeft + iSignalRight) \ 2
        For iKey = 0 To NumberOfNotes
            KeyPowerReal(iKey) = KeyPowerReal(iKey) + iSignal * (ValueCos(iKey, 2))
            KeyPowerImag(iKey) = KeyPowerImag(iKey) + iSignal * (ValueSin(iKey, 2))
        Next
        
        For iKey = 0 To NumberOfNotes
            vCos = ValueCos(iKey, 2) * CosKey(iKey) - ValueSin(iKey, 2) * SinKey(iKey)
            vSin = ValueSin(iKey, 2) * CosKey(iKey) + ValueCos(iKey, 2) * SinKey(iKey)
            ValueCos(iKey, 2) = vCos
            ValueSin(iKey, 2) = vSin
        Next
    Next
    
    While Playing
        For iFileNum = 1 To 2

            Get #iFileNum, , iSignalLeft
            Get #iFileNum, , iSignalRight
            iSignal = (iSignalLeft + iSignalRight) \ 2
            
            iDirection = (iFileNum * 2) - 3
            iSignal = iSignal * iDirection
            
            'lSample(iFileNum) = lSample(iFileNum) + 1
            
            For iKey = 0 To NumberOfNotes
                KeyPowerReal(iKey) = KeyPowerReal(iKey) + iSignal * (ValueCos(iKey, iFileNum))
                KeyPowerImag(iKey) = KeyPowerImag(iKey) + iSignal * (ValueSin(iKey, iFileNum))
            Next
            
            For iKey = 0 To NumberOfNotes
                vCos = ValueCos(iKey, iFileNum)
                ValueCos(iKey, iFileNum) = ValueCos(iKey, iFileNum) * CosKey(iKey) - ValueSin(iKey, iFileNum) * SinKey(iKey)
                ValueSin(iKey, iFileNum) = ValueSin(iKey, iFileNum) * CosKey(iKey) + vCos * SinKey(iKey)
            Next
        Next
        
        iShowKeyboard = iShowKeyboard - 1
        If iShowKeyboard < 1 Then
            DisplayFretBoard (iSamples)
            txtTime = Format((((Seek(1) - DataStart) \ 4) / SampleRate), ".000")
            DoEvents
            iShowKeyboard = CLng(SampleRate) * Val(txtStep.Text)
            iSamples = iSamples + 1
        End If
    Wend
End Sub

Private Sub DisplayFretBoard(iTime As Long)
    Dim scan As Long
    Dim dPowerReal As Double
    Dim dPowerImag As Double
    Dim power As Double
    Dim red As Long
    Dim green As Long
    Dim blue As Long
    Dim keycolour As Long
    Dim phase As Double
    Dim intensity As Double
    Dim offset As Long
    
    ' Work out the power for each key
    For scan = 0 To NumberOfNotes Step 1
        dPowerReal = 0
        dPowerImag = 0
        
        For offset = -1 To 1
            dPowerReal = dPowerReal + KeyPowerReal(scan + offset) / DFTSize
            dPowerImag = dPowerImag + KeyPowerImag(scan + offset) / DFTSize
        Next

        power = Sqr(dPowerReal * dPowerReal + dPowerImag * dPowerImag)
        phase = Angle(dPowerReal, dPowerImag)
        
       ' intensity = phase
        intensity = 1 - Exp(-power / 80)
        
        keycolour = vbWhite
        red = keycolour And vbRed
        green = (keycolour And vbGreen) \ &H100
        blue = (keycolour And vbBlue) \ &H10000

        DrawNote scan, RGB(red * intensity, green * intensity, blue * intensity), iTime
    Next
End Sub

Private Sub DrawNote(ByVal iNote As Integer, ByVal lColour As Long, ByVal iTime As Long)
    Dim iString As Integer
    Dim iOffset As Integer
    Dim iPos As Integer
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim lPos As Long
    
    Display.Line (iNote * NoteWidth, iTime)-Step(NoteWidth, 0), lColour
    
End Sub

Private Function Angle(dx As Double, dy As Double) As Double
    Dim sgx As Long
    Dim sgy As Long
    Dim quad As Long
    
    If dy = 0 Then
        If dx > 0 Then
            Angle = 0
        Else
            Angle = 0.5
        End If
        Exit Function
    End If
    
    If dx = 0 Then
        If dy > 0 Then
            Angle = 0.25
        Else
            Angle = 0.75
        End If
        Exit Function
    End If
    
    quad = (1 - Sgn(dx)) \ 2 + (1 - Sgn(dy))
    Select Case quad
        Case 0
            Angle = Atn(dy / dx) / (pi * 2)
        Case 1
            Angle = Atn(dy / dx) / (pi * 2) + 0.5
        Case 3
            Angle = Atn(dy / dx) / (pi * 2) + 0.5
        Case 2
            Angle = Atn(dy / dx) / (pi * 2) + 1
    End Select
End Function
