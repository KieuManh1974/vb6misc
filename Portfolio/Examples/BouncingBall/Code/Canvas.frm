VERSION 5.00
Begin VB.Form Canvas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Executive Bouncing Ball"
   ClientHeight    =   8010
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7560
   DrawMode        =   7  'Invert
   FillColor       =   &H000000FF&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   12
      Left            =   1080
      Top             =   1200
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuNoGravity 
         Caption         =   "&No Gravity"
      End
      Begin VB.Menu mnuGravity 
         Caption         =   "&Gravity"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRubberBall 
         Caption         =   "Rubber &Ball"
      End
      Begin VB.Menu mnuLeadBall 
         Caption         =   "&Lead Ball"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFreeze 
         Caption         =   "&Freeze"
      End
      Begin VB.Menu mnuUnfreeze 
         Caption         =   "&Unfreeze"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetBallBack 
         Caption         =   "&Get Ball Back"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Balls As New Collection
Public Walls As New Collection


Private Sub Form_Load()
    DoEvents
    DoEvents
    Initialise
    'Test
End Sub

Private Sub Initialise()
    Dim AWall As Wall
    Set AWall = New Wall
    Set AWall.Canvas = Me
    AWall.DrawWall 0, Me.ScaleHeight - 40, Me.ScaleWidth - 40, Me.ScaleHeight
    Walls.Add AWall

    Set AWall = New Wall
    Set AWall.Canvas = Me
    AWall.DrawWall 0, 0, 40, Me.ScaleHeight - 40
    Walls.Add AWall

    Set AWall = New Wall
    Set AWall.Canvas = Me
    AWall.DrawWall Me.ScaleWidth - 40, 0, Me.ScaleWidth, Me.ScaleHeight
    Walls.Add AWall

    Set AWall = New Wall
    Set AWall.Canvas = Me
    AWall.DrawWall 0, 0, Me.ScaleWidth, 40
    Walls.Add AWall

    Dim P As Long
    For P = 1 To 5
        Set AWall = New Wall
        Set AWall.Canvas = Me
        AWall.DrawPartition P * 50 + 100, 100, P * 50 + 100, 300
        Walls.Add AWall
    Next
        
'    Set AWall = New Wall
'    Set AWall.Canvas = Me
'    AWall.Polygon 300, 350, 150, 20
'    Walls.Add AWall
'
'    Set AWall = New Wall
'    Set AWall.Canvas = Me
'    AWall.Polygon 550, 450, 100, 4
'    Walls.Add AWall
    
        
'
'    Dim ABall As Ball
'
'    Dim q As Integer
'    Randomize
'    For q = 1 To 2
'        Set ABall = New Ball
'        ABall.Size = 30
'        ABall.Initialise Rnd * 300, Rnd * 400, Rnd * 100 - 50, Rnd * 100 - 50
'        ABall.Interval = Timer.Interval / 1000
'        Set ABall.Parent = Me
'        Balls.Add ABall
'    Next
    
    Set ABall = New Ball
    ABall.Size = 15
    ABall.Initialise 100, 470 - 100, 50, 0
    ABall.Interval = Timer.Interval / 1000
    Set ABall.Parent = Me
    Balls.Add ABall
        
    Timer.Enabled = True
End Sub

' Check to see if the mouse has picked anything up
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim vBall As Ball
    Dim vWall As Wall
    Dim vSide As Variant
    
    Dim dist As Double
    Dim bPickedUpBall As Boolean
    
    ' See if we've picked up any balls
    For Each vBall In Balls
        With vBall
            If Button = vbLeftButton Then
                If .Frozen Then
                    .SetPostion CDbl(X), CDbl(Y)
                    .DrawBall
                    .RemoveOldBall
                    .SetVelocity (.BallX - .OldX) * 40, (.BallY - .OldY) * 40
                    .SaveToOldAnchor
                    bPickedUpBall = True
                Else
                    dist = (X - .BallX) ^ 2 + (Y - (.BallY)) ^ 2
                    If dist < (.Size * .Size) Then
                        .Initialise CDbl(X), CDbl(Y), 0, 0
                        .Frozen = True
                        bPickedUpBall = True
                    End If
                End If
            End If
        End With
    Next
    
    If bPickedUpBall Then
        Exit Sub
    End If
    
    ' See if we've picked up the end of a wall
    Dim iSideIndex As Integer
    Dim iWallIndex As Integer
    
    For iWallIndex = 1 To Walls.Count
        Set vWall = Walls(iWallIndex)
        For iSideIndex = 1 To vWall.SideCount
            If Button = vbLeftButton Then
                Select Case vWall.Captured(iSideIndex)
                    Case 1
                        Me.Line (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1))-(vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1)), 2, vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), 2, vbGreen
                        
                        Walls(iWallIndex).SideValue(iSideIndex, 0) = X
                        Walls(iWallIndex).SideValue(iSideIndex, 1) = Y
                        Me.Line (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1))-(vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1)), 2, vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), 2, vbGreen
                        If Shift = 1 Then
                            Exit Sub
                        End If
                    Case 2
                        Me.Line (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1))-(vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1)), 2, vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), 2, vbGreen
                        Walls(iWallIndex).SideValue(iSideIndex, 2) = X
                        Walls(iWallIndex).SideValue(iSideIndex, 3) = Y
                        Me.Line (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1))-(vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(0), vWall.Side(iSideIndex)(1)), 2, vbGreen
                        Me.Circle (vWall.Side(iSideIndex)(2), vWall.Side(iSideIndex)(3)), 2, vbGreen
                        If Shift = 1 Then
                            Exit Sub
                        End If
                    Case Else
                        dist1 = (X - vWall.Side(iSideIndex)(0)) ^ 2 + (Y - vWall.Side(iSideIndex)(1)) ^ 2
                        dist2 = (X - vWall.Side(iSideIndex)(2)) ^ 2 + (Y - vWall.Side(iSideIndex)(3)) ^ 2
                        If dist1 < 50 Then
                            vWall.Captured(iSideIndex) = 1
                            If Shift = 1 Then
                                Exit Sub
                            End If
                        End If
                        If dist2 < 50 Then
                            vWall.Captured(iSideIndex) = 2
                            If Shift = 1 Then
                                Exit Sub
                            End If
                        End If
                End Select
            End If
        Next
    Next
End Sub

' Let go of the ball or wall
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim vBall As Ball
    Dim vWall As Wall
    
    For Each vBall In Balls
        With vBall
            If .Frozen Then
                .Frozen = False
            End If
        End With
    Next
    
    For Each vWall In Walls
        For iSideIndex = 1 To vWall.SideCount
            vWall.Captured(iSideIndex) = 0
        Next
    Next
End Sub


Private Sub mnuFreeze_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Frozen = True
        End With
    Next
End Sub

Private Sub mnuGetBallBack_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            BallSeg.Position CDbl(200), CDbl(200)
            .Frozen = True
        End With
    Next
End Sub

Private Sub mnuGravity_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Gravity = -800
        End With
    Next
    mnuNoGravity.Checked = False
    mnuGravity.Checked = True
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show vbModal
End Sub

Private Sub mnuLeadBall_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Resistance = 0.4
        End With
    Next
    mnuLeadBall.Checked = True
    mnuRubberBall.Checked = False
End Sub

Private Sub mnuNoGravity_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Gravity = 0
        End With
    Next
    mnuNoGravity.Checked = True
    mnuGravity.Checked = False
End Sub



Private Sub mnuRubberBall_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Resistance = 0#
        End With
    Next
    mnuLeadBall.Checked = False
    mnuRubberBall.Checked = True
End Sub

Private Sub mnuUnfreeze_Click()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            .Frozen = False
        End With
    Next
End Sub

Private Sub Timer_Timer()
    Dim vBall As Ball
    
    For Each vBall In Balls
        With vBall
            If .OldX <> 0 And .OldY <> 0 Then
                .RemoveOldBall
            End If
            .DrawBall
            .SaveToOldAnchor
            .SaveToOldDirection
        End With
    Next
        
    For Each vBall In Balls
        vBall.UpdateBall
        vBall.UpdateWallCollisions
    Next
    
'    For Each vBall In Balls
'        vBall.SaveFromNewSeg
'    Next
    
'    For Each vBall In Balls
'        vBall.UpdateBallCollisions
'    Next
    
'    For Each vBall In Balls
'        vBall.SaveFromNewSeg
'    Next
End Sub
