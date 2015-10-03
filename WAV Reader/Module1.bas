Attribute VB_Name = "Module1"
Option Explicit

Private Const DFTSize = 65536

Private Const pi = 3.14159265358979

'Private Const NumberOfNotes = 88 - 1
Private Const NumberOfNotes = 700
Private Const NoteDivision = 96
Private Const NoteWidth = 1
Private DataStart As Long
Private SampleRate As Long

Private Type Sample
    SignalLeft As Integer
    SignalRight As Integer
End Type

Private StreamIn() As Sample
Private StreamOut() As Sample

Private Oscillator(1 To 1300, 1300) As Long
Private Index(1 To 1300) As Long
Private StreamLength As Long
Private Const StreamBuffer As Long = 10

Dim sFile As String
Dim sTargetFile As String

Sub main()
    sFile = "C:\Main\Programming Projects\Projects\Miscellaneous\WAV Reader\Track1.wav"
    sTargetFile = "C:\Main\Programming Projects\Projects\Miscellaneous\WAV Reader\compressed.wav"
    LoadFile
    CompressFile
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
    
    If Dir(sFile) <> "" Then
        ' Open the leading file
        Close #1
        Open sFile For Binary As #1
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
        StreamLength = subchunk2size
        ReDim StreamIn(StreamBuffer) As Sample
        Get #1, , StreamIn
        
        DataStart = Seek(1)
        
        ' Open the following file
        Close #2
        Open sTargetFile For Binary As #2
        Put #2, , chunkdescriptor
        Put #2, , chunksize
        Put #2, , wave
        Put #2, , fmt
        Put #2, , subchunk1size
        Put #2, , audioformat
        Put #2, , numchanels
        Put #2, , SampleRate
        Put #2, , byterate
        Put #2, , blockalign
        Put #2, , bitspersample
        Put #2, , data
        Put #2, , subchunk2size
        Seek #2, DataStart
    End If
    
End Sub

Private Sub CompressFile()
    Dim iSignalLeft As Integer
    Dim iSignalRight As Integer
    Dim x As Long
    Dim sum As Double
    Dim lSampleIndex   As Long
    Dim StreamIndex As Long
    
    ReDim StreamOut(StreamLength) As Sample
    
    For lSampleIndex = 0 To StreamLength / 4
        Get #1, , StreamIn(StreamIndex): StreamIndex = StreamIndex + 1: If StreamIndex >= StreamBuffer Then StreamIndex = 0

        If Rnd < 0.5 Then
            Put #2, , StreamIn(StreamIndex).SignalLeft * 2
            Put #2, , CInt(Rnd * 1000)
        Else
            Put #2, , CInt(Rnd * 1000)
            Put #2, , StreamIn(StreamIndex).SignalRight * 2
        End If
        
    Next
    
    Close #1
    Close #2
End Sub

Private Function Mult(x As Integer, y As Integer) As Integer
    If CDbl(x) * y > 32767 Then
        Mult = 32767
    ElseIf CDbl(x) * y < -32768 Then
        Mult = -32768
    Else
        Mult = x * y
    End If
End Function


