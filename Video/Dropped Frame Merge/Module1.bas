Attribute VB_Name = "Module1"
Option Explicit

Private DropFrames() As Boolean

Private Const TOTAL_FRAMES = 537
Private Const FRAME_SETS = 3

Sub main()
    Dim lFrameSetIndex As Long
    Dim lFrameIndex As Long
    Dim bDropped  As Boolean
    
    ReDim DropFrames(FRAME_SETS, TOTAL_FRAMES) As Boolean
    
    For lFrameSetIndex = 0 To FRAME_SETS
        ScanFrames lFrameSetIndex
    Next
    
    ' Amalgamate
    For lFrameIndex = 0 To TOTAL_FRAMES
        bDropped = True
        For lFrameSetIndex = 0 To FRAME_SETS
            If Not DropFrames(lFrameSetIndex, lFrameIndex) Then
                bDropped = False
            End If
        Next
        If bDropped Then
            Debug.Print lFrameIndex & " dropped"
        End If
    Next
End Sub

Public Sub ScanFrames(ByVal lFrameset As Long)
    Dim lCount As Long
    Dim f1() As Byte
    Dim f2() As Byte
    
    Dim lCompare As Long
    Dim bDropped As Boolean
    
    For lCount = 47 To TOTAL_FRAMES
        Open "D:\Media\Video\Captured\Frames" & lFrameset & "\FF." & Format$(lCount - 1, "00") & ".avi" For Binary As #1
        Open "D:\Media\Video\Captured\Frames" & lFrameset & "\FF." & Format$(lCount, "00") & ".avi" For Binary As #2
        
        bDropped = True
        If LOF(1) = LOF(2) Then
            ReDim f1(LOF(1)) As Byte
            ReDim f2(LOF(2)) As Byte
            
            Get #1, , f1()
            Get #2, , f2()
                
            For lCompare = 0 To UBound(f1)
                If f1(lCompare) <> f2(lCompare) Then
                    bDropped = False
                    Exit For
                End If
            Next
        Else
            bDropped = False
        End If
        
        Close #1
        Close #2
        
        DropFrames(lFrameset, lCount) = bDropped
        If bDropped Then
            Debug.Print lFrameset & "." & lCount
        End If
    Next
End Sub
