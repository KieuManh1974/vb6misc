Attribute VB_Name = "Functions"
Option Explicit


Public Function LineThroughCircle(Q As Vector, v As Vector, P As Vector, dRadius As Double) As Variant
    Dim Roots As Variant
    Dim S1 As Vector
    Dim S2 As Vector
    
    Dim n1 As Double
    Dim n2 As Double
    Dim n3 As Double
    Dim D As Vector
    Dim P1 As Vector
    Dim P2 As Vector
    Dim sv As Double
    
    Set D = Subtract(Q, P)
    n1 = Dot(v, v)
    n2 = 2 * Dot(v, D)
    n3 = Dot(D, D) - dRadius * dRadius
    
    Roots = Quadratic(n1, n2, n3)

    If UBound(Roots) > 0 Then
        Set P1 = Add(Q, Scalar(v, Roots(0)))
        Set P2 = Add(Q, Scalar(v, Roots(1)))
        
        sv = SizeOf(v)
        If Roots(0) < Roots(1) Then
            LineThroughCircle = Array(Roots(0) * sv, Roots(1) * sv, P1, P2)
        Else
            LineThroughCircle = Array(Roots(1) * sv, Roots(0) * sv, P2, P1)
        End If
    Else
        LineThroughCircle = Array()
    End If
End Function

Public Function Quadratic(a As Double, b As Double, c As Double) As Variant
    Dim determinant As Double
    
    determinant = b ^ 2 - 4 * a * c
    If determinant < 0 Or a = 0 Then
        Quadratic = Array()
        Exit Function
    End If
    
    Quadratic = Array((-b + Sqr(determinant)) / (2 * a), (-b - Sqr(determinant)) / (2 * a))
End Function

Public Sub Test()
    Dim vRoots As Variant
    
    vRoots = LineThroughCircle(VectorOf(0, 0), VectorOf(230, 250), VectorOf(100, 100), 30)
    If UBound(vRoots) > -1 Then
        Canvas.Circle (vRoots(2).X, vRoots(2).Y), 5
        'Canvas.Circle (vRoots(3).X, vRoots(3).Y), 5
    End If
    Canvas.Circle (100, 100), 30
    Canvas.Line (0, 0)-(230, 250)
    
End Sub


Public Function Crosses(a As VectorSet, b As VectorSet) As Boolean
    Dim Side1 As Integer
    Dim Side2 As Integer
    Dim Side3 As Integer
    Dim Side4 As Integer
    
    Side1 = Sgn(Cross(a.Velocity, Subtract(b.Position, a.Position)))
    Side2 = Sgn(Cross(a.Velocity, Subtract(Add(b.Position, b.Velocity), a.Position)))
    Side3 = Sgn(Cross(b.Velocity, Subtract(a.Position, b.Position)))
    Side4 = Sgn(Cross(b.Velocity, Subtract(Add(a.Position, a.Velocity), b.Position)))
    
    If (Side1 * Side2) < 0 And (Side3 * Side4) < 0 Then
        Crosses = True
    End If
    
End Function

Public Function DistanceToIntersection(b As VectorSet, w As VectorSet) As Double
    Dim Numerator As Double
    Dim Denominator As Double
    Dim D As Vector
    
    Set D = Subtract(b.Position, w.Position)
    Numerator = Cross(w.Velocity, D)
    Denominator = Cross(b.Velocity, w.Velocity)
    
    If Denominator <> 0 Then
        DistanceToIntersection = Numerator / Denominator
    Else
        DistanceToIntersection = 1000000# 'Large value
    End If
End Function

Public Function Reflect(incident As Vector, plane As Vector) As Vector
    Dim rotate As Vector
    Dim rotateback As Vector
    Dim incidentnorm As Vector
    
    Set rotate = Scalar(VectorOf(plane.X, -plane.Y), 1 / SizeOf(plane))
    Set rotateback = Scalar(plane, 1 / SizeOf(plane))
    Set incidentnorm = Multiply(incident, rotate)
    Set incidentnorm = VectorOf(incidentnorm.X, -incidentnorm.Y)
    Set Reflect = Multiply(incidentnorm, rotateback)
End Function

Public Function VectorOf(ByVal X As Double, ByVal Y As Double) As Vector
    Set VectorOf = New Vector
    VectorOf.X = X
    VectorOf.Y = Y
End Function

Public Function SegmentOf(a As Vector, b As Vector) As VectorSet
    Set SegmentOf = New VectorSet
    Set SegmentOf.Position = a
    Set SegmentOf.Velocity = b
End Function

Public Function DynamicsOf(a As Vector, b As Vector, c As Vector) As VectorSet
    Set DynamicsOf = New VectorSet
    Set DynamicsOf.Position = a
    Set DynamicsOf.Velocity = b
    Set DynamicsOf.Acceleration = c
End Function

Public Function Scalar(v1 As Vector, ByVal factor As Double) As Vector
    Set Scalar = New Vector
    Scalar.X = v1.X * factor
    Scalar.Y = v1.Y * factor
End Function

Public Function Multiply(v1 As Vector, v2 As Vector) As Vector
    Set Multiply = New Vector
    Multiply.X = v1.X * v2.X - v1.Y * v2.Y
    Multiply.Y = v1.Y * v2.X + v1.X * v2.Y
End Function

Public Function Add(v1 As Vector, v2 As Vector) As Vector
    Set Add = New Vector
    Add.X = v1.X + v2.X
    Add.Y = v1.Y + v2.Y
End Function

Public Function Subtract(v1 As Vector, v2 As Vector) As Vector
    Set Subtract = New Vector
    Subtract.X = v1.X - v2.X
    Subtract.Y = v1.Y - v2.Y
End Function

Public Function Distance(v1 As Vector, v2 As Vector) As Double
    Distance = Sqr((v2.X - v1.X) * (v2.X - v1.X) + (v2.Y - v1.Y) * (v2.Y - v1.Y))
End Function

Public Function SizeOf(v As Vector) As Double
    SizeOf = Sqr(v.X * v.X + v.Y * v.Y)
End Function

Public Function Cross(v1 As Vector, v2 As Vector) As Double
    Cross = v1.X * v2.Y - v1.Y * v2.X
End Function

Public Function Dot(v1 As Vector, v2 As Vector) As Double
    Dot = v1.X * v2.X + v1.Y * v2.Y
End Function

Public Function Binomial(n1 As Double, n2 As Double) As Variant
    Binomial = Array(n1 * n1, 2 * n1 * n2, n2 * n2)
End Function

Public Function CopyVector(v As Vector) As Vector
    Set CopyVector = New Vector
    CopyVector.X = v.X
    CopyVector.Y = v.Y
End Function



