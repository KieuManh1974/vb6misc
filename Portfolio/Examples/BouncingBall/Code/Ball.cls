VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private BallSeg As New VectorSet
Private OldSeg As New VectorSet
Private NewSeg As New VectorSet

Private BallDynamics As VectorSet

Public BallMass As Double
Public Size As Double
Public Interval As Double

Public Frozen As Boolean

Public Parent As Canvas

Public Resistance As Double
Public Gravity As Double

Private Sub Class_Initialize()
    Gravity = -800
    Resistance = 0.4
End Sub


Public Property Get BallX() As Double
    BallX = BallSeg.Position.X
End Property

Public Property Get BallY() As Double
    BallY = BallSeg.Position.Y
End Property

Public Property Get BallDX() As Double
    BallDX = BallSeg.Velocity.X
End Property

Public Property Get BallDY() As Double
    BallDY = BallSeg.Velocity.Y
End Property

Public Property Get OldX() As Double
    OldX = OldSeg.Position.X
End Property

Public Property Get OldY() As Double
    OldY = OldSeg.Position.Y
End Property

Public Sub Initialise(X As Double, Y As Double, DX As Double, DY As Double)
    BallSeg.Position.X = X
    BallSeg.Position.Y = Y
    BallSeg.Velocity.X = DX
    BallSeg.Velocity.Y = DY
End Sub
        
Public Sub SaveToOldAnchor()
    Set OldSeg.Position = CopyVector(BallSeg.Position)
End Sub

Public Sub SaveToOldDirection()
    Set OldSeg.Velocity = CopyVector(BallSeg.Velocity)
End Sub

Public Sub SaveFromNewSeg()
    Set BallSeg = CopyVector(NewSeg)
End Sub

Public Sub SetPostion(X As Double, Y As Double)
    BallSeg.Position.X = X
    BallSeg.Position.Y = Y
End Sub

Public Sub SetVelocity(X As Double, Y As Double)
    BallSeg.Velocity.X = X
    BallSeg.Velocity.Y = Y
End Sub

Public Sub UpdateBall()
    If Frozen Then
        Exit Sub
    End If
    
    Set BallSeg.Position = Add(BallSeg.Position, Scalar(BallSeg.Velocity, Interval))
    BallSeg.Velocity.Y = BallSeg.Velocity.Y - Gravity * Interval
End Sub

Public Sub UpdateBall2()
    If Frozen Then
        Exit Sub
    End If
    
    BallDynamics.Velocity.X = BallDynamics.Velocity.X + BallDynamics.Acceleration.X * Interval
    BallDynamics.Velocity.Y = BallDynamics.Velocity.Y + BallDynamics.Acceleration.Y * Interval
    
    BallDynamics.Position.X = BallDynamics.Position.X + BallDynamics.Velocity.X
    BallDynamics.Position.Y = BallDynamics.Position.Y + BallDynamics.Velocity.Y
    
End Sub

Public Sub UpdateWallCollisions()
    'DetectWallCollisions
    Set NewSeg = DetectWallCollisions2(BallSeg, Size)
End Sub

Public Sub UpdateBallCollisions()
    DetectBallCollisions
End Sub

Public Sub DrawBall()
    Parent.Circle (BallSeg.Position.X, BallSeg.Position.Y), Size
    Parent.Circle (BallSeg.Position.X, BallSeg.Position.Y), 2
End Sub

Public Sub RemoveOldBall()
    Parent.Circle (OldSeg.Position.X, OldSeg.Position.Y), Size
    Parent.Circle (OldSeg.Position.X, OldSeg.Position.Y), 2
End Sub

Private Function DetectWallCollisions2(CurrentBall As VectorSet, ByVal CurrentSize As Double) As VectorSet
    Dim vWall As Wall
    Dim vSide As Variant
    Dim Reflection As Vector
    Dim BallSegOffset As VectorSet
    Dim WallSeg As VectorSet
    Dim PreDistance As Double
    Dim ActualDistance As Double
    Dim BallVelocity As Double
    Dim Perpendicular As Vector
    Dim WallDirection As Integer
    Dim CrossPoint As Vector
    Dim AngleOfResistance As Double
    Dim iSideIndex As Integer

    Dim P1 As Vector
    Dim P2 As Vector
    Dim W1 As Vector
    Dim W2 As Vector
    Dim R1 As Vector
    Dim R2 As Vector
    Dim WallLeftRight As Double
    Dim OffsetLine As VectorSet
    Dim DistanceToCollision As Double
    
    Dim D1 As Vector
    Dim D2 As Vector
    Dim D1Size As Double
    Dim R1Dist As Double
    Dim iWallIndex As Long
    Dim cs As Double
    Dim sn As Double
    Dim CircleIntersects As Variant
    
    Set DetectWallCollisions2 = CurrentBall
    
    For iWallIndex = 1 To Parent.Walls.Count
        Set vWall = Parent.Walls(iWallIndex)
        For iSideIndex = 1 To vWall.SideCount
            vSide = vWall.Side(iSideIndex)
            Set P1 = VectorOf(vSide(0), vSide(1))
            Set P2 = VectorOf(vSide(2), vSide(3))
            Set WallSeg = SegmentOf(P1, Subtract(P2, P1))
            WallLeftRight = -Sgn(Cross(Subtract(CurrentBall.Position, WallSeg.Position), WallSeg.Velocity))
            Set Perpendicular = Scalar(Multiply(WallSeg.Velocity, VectorOf(0, WallLeftRight)), CurrentSize / SizeOf(WallSeg.Velocity))
            Set W1 = Add(P1, Perpendicular)
            Set W2 = Add(P2, Perpendicular)
            Set OffsetLine = SegmentOf(W1, Subtract(W2, W1))

            If Crosses(SegmentOf(CurrentBall.Position, Scalar(CurrentBall.Velocity, Interval)), OffsetLine) Then
                DistanceToCollision = DistanceToIntersection(CurrentBall, WallSeg)
                Set Reflection = Reflect(CurrentBall.Velocity, WallSeg.Velocity)
                Set DetectWallCollisions2.Velocity = Reflection
                
                'Parent.Line (R1.X, R1.Y)-(R2.X, R2.Y)
            Else ' corner detection
                CircleIntersects = LineThroughCircle(CurrentBall.Position, CurrentBall.Velocity, P1, CurrentSize)
                If UBound(CircleIntersects) > -1 Then ' Possible Intersection
                    If CircleIntersects(0) > 0 Then
                        If SizeOf(CurrentBall.Velocity) * Interval >= CircleIntersects(0) Then
                            Set DetectWallCollisions2.Velocity = Scalar(Reflect(CurrentBall.Velocity, Subtract(P1, CurrentBall.Position)), -1)
                        End If
                    End If
                End If

                CircleIntersects = LineThroughCircle(CurrentBall.Position, CurrentBall.Velocity, P2, CurrentSize)
                If UBound(CircleIntersects) > -1 Then ' Possible Intersection
                    If CircleIntersects(0) > 0 Then
                        If SizeOf(CurrentBall.Velocity) * Interval >= CircleIntersects(0) Then
                            Set DetectWallCollisions2.Velocity = Scalar(Reflect(CurrentBall.Velocity, Subtract(P2, CurrentBall.Position)), -1)
                        End If
                    End If
                End If

            End If
        Next
    Next
End Function

Private Sub DetectWallCollisions()
    Dim vWall As Wall
    Dim vSide As Variant
    Dim Reflection As Vector
    Dim BallSegOffset As VectorSet
    Dim WallSeg As VectorSet
    Dim PreDistance As Double
    Dim ActualDistance As Double
    Dim BallVelocity As Double
    Dim Perpendicular As Vector
    Dim WallDirection As Integer
    Dim CrossPoint As Vector
    Dim AngleOfResistance As Double
    Dim iSideIndex As Integer
    
    NewSeg = BallSeg
    
    BallVelocity = SizeOf(BallSeg.Velocity)

    For Each vWall In Parent.Walls
        For iSideIndex = 1 To vWall.SideCount
            vSide = vWall.Side(iSideIndex)
            WallSeg = SegmentOf(VectorOf(vSide(0), vSide(1)), Subtract(VectorOf(vSide(2), vSide(3)), VectorOf(vSide(0), vSide(1))))
            WallDirection = Sgn(Cross(Subtract(BallSeg.Position, WallSeg.Position), WallSeg.Velocity))
            Perpendicular = Scalar(Multiply(WallSeg.Velocity, VectorOf(0, WallDirection)), Size / SizeOf(WallSeg.Velocity))
            BallSegOffset = SegmentOf(Add(BallSeg.Position, Perpendicular), BallSeg.Velocity)
            
             If Crosses(SegmentOf(BallSegOffset.Position, Scalar(BallSegOffset.Velocity, Interval)), WallSeg) Then
                PreDistance = DistanceToIntersection(BallSegOffset, WallSeg)
                ActualDistance = BallVelocity * PreDistance
                If ActualDistance < (BallVelocity * Interval) And ActualDistance > 0 Then  ' Will collide in next frame
                        Reflection = Reflect(BallSegOffset.Velocity, WallSeg.Velocity)
                        AngleOfResistance = Resistance * (1 - Abs(Dot(BallSeg.Velocity, WallSeg.Velocity) / (SizeOf(BallSeg.Velocity) * SizeOf(WallSeg.Velocity))))
'                        BallSeg.Velocity.X = Reflection.X - AngleOfResistance * Reflection.X
'                        BallSeg.Velocity.Y = Reflection.Y - AngleOfResistance * Reflection.Y
                              
                        NewSeg.Position = BallSeg.Position
                        NewSeg.Velocity.X = Reflection.X - AngleOfResistance * Reflection.X
                        NewSeg.Velocity.Y = Reflection.Y - AngleOfResistance * Reflection.Y
                        
                        ' Work out cross point
                        'CrossPoint = Add(Scalar(BallSegOffset.Velocity, ActualDistance / Size), BallSegOffset.Position)
                        'ballseg.Position = add(crosspoint,neg()
                        Exit Sub
                End If
            End If
        Next
    Next
End Sub

Private Sub DetectBallCollisions()
    Dim vBall As Ball
    Dim vBallSeg As VectorSet
    Dim bin As Variant
    Dim sol As Variant
    Dim DiffConst As Vector
    Dim DiffT As Vector
    Dim LeastTime As Double
    Dim quad As Variant
    Dim dist As Double
    Dim binx As Variant
    Dim biny As Variant
    
    NewSeg = BallSeg
    
    For Each vBall In Parent.Balls
        With vBall
            dist = Size + vBall.Size
            
            vBallSeg = SegmentOf(VectorOf(vBall.BallX, vBall.BallY), VectorOf(vBall.BallDX, vBall.BallDY))
            
            DiffConst = Subtract(vBallSeg.Position, BallSeg.Position)
            DiffT = Subtract(vBallSeg.Velocity, BallSeg.Velocity)
            binx = Binomial(DiffT.X, DiffConst.X)
            biny = Binomial(DiffT.Y, DiffConst.Y)
            quad = Array(binx(0) + biny(0), binx(1) + biny(1), binx(2) + biny(2) - dist * dist)
            sol = Quadratic(CDbl(quad(0)), CDbl(quad(1)), CDbl(quad(2)))
            
            If Not IsNull(sol) Then
                If sol(0) > 0 And sol(0) < Interval Then
                    If sol(1) < sol(0) And sol(1) > 0 Then
                        LeastTime = sol(1)
                        'Beep
                    Else
                        LeastTime = sol(0)
                        'Beep
                    End If
                Else
                    If sol(1) > 0 And sol(1) < Interval Then
                        LeastTime = sol(1)
                        'Beep
                    Else
                        Exit Sub
                    End If
                End If
            Else
                Exit Sub
            End If
            
            Dim MePos As Vector
            Dim OtherPos As Vector
            Dim MePerp As Vector
            Dim OtherPerp As Vector
            
            ' Project both motions forward
            'MePos = Add(BallSeg.Position, Scalar(BallSeg.Velocity, LeastTime))
            'OtherPos = Add(vBallSeg.Position, Scalar(vBallSeg.Velocity, LeastTime))
            
            OtherPerp = Multiply(vBallSeg.Velocity, VectorOf(0, -1))
            
            ' Work out new trajectories
            NewSeg.Position = BallSeg.Position
            NewSeg.Velocity = Reflect(BallSeg.Velocity, OtherPerp)
            
        End With
    Next
End Sub
