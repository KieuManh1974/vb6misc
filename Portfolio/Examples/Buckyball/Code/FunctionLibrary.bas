Attribute VB_Name = "FunctionLibrary"
Option Explicit

Public Type Vector
    x As Single
    y As Single
    z As Single
End Type

Public Type Vector2d
    x As Single
    y As Single
End Type

Public Type Orientation
    axial As Vector
    rightside As Vector
    down As Vector
    position As Vector
End Type

Public Type LineSet
    start As Integer
    finish As Integer
    visible As Boolean
End Type

Public Type VectorSet
    lineref() As Integer
    visible As Boolean
End Type

Public Type Entity
    positional As Orientation
    movement As Orientation
    points() As Vector
    points2d() As Vector2d
    lines() As LineSet
    facets() As VectorSet
End Type

Public Function ConvertTo2d(position As Vector) As Vector2d
    ConvertTo2d.x = 20000 * position.x / (2500 - position.z)
    ConvertTo2d.y = 20000 * position.y / (2500 - position.z)
End Function

Public Function VectorOf(x As Double, y As Double, z As Double) As Vector
    VectorOf.x = x
    VectorOf.y = y
    VectorOf.z = z
End Function

Public Function Add(a As Vector, b As Vector) As Vector
    Add.x = a.x + b.x
    Add.y = a.y + b.y
    Add.z = a.z + b.z
End Function

Public Function Add2(a As Vector, b As Vector, c As Vector) As Vector
    Add2.x = a.x + b.x + c.x
    Add2.y = a.y + b.y + c.y
    Add2.z = a.z + b.z + c.z
End Function

Public Function Subs(a As Vector, b As Vector) As Vector
    Subs.x = a.x - b.x
    Subs.y = a.y - b.y
    Subs.z = a.z - b.z
End Function

Public Function Sub2d(a As Vector2d, b As Vector2d) As Vector2d
    Sub2d.x = a.x - b.x
    Sub2d.y = a.y - b.y
End Function

Public Function Scalar(a As Vector, ByVal size As Single) As Vector
    Scalar.x = a.x * size
    Scalar.y = a.y * size
    Scalar.z = a.z * size
End Function

Public Function Dot(a As Vector, b As Vector) As Double
    Dot = a.x * b.x + a.y * b.y + a.z * b.z
End Function

Public Function Cross(a As Vector, b As Vector) As Vector
    Cross.x = a.y * b.z - b.y * a.z
    Cross.y = a.z * b.x - b.z * a.x
    Cross.z = a.x * b.y - a.y * b.x
End Function

Public Function Cross2d(a As Vector2d, b As Vector2d) As Double
    Cross2d = a.x * b.y - a.y * b.x
End Function

Public Function Mag(a As Vector) As Double
    Mag = Sqr(a.x * a.x + a.y * a.y + a.z * a.z)
End Function

Public Function JoinOf(start As Integer, finish As Integer) As LineSet
    JoinOf.start = start
    JoinOf.finish = finish
End Function

Public Function Grid(position As Vector, xVector As Vector, yVector As Vector, zVector As Vector) As Vector
    Grid = Add2(Scalar(xVector, position.x), Scalar(yVector, position.y), Scalar(zVector, position.z))
End Function

Public Function Rotate(thevector As Vector, ByVal angle As Single, Axis As Vector)
    Dim origin As Vector
    Dim xaxis As Vector
    Dim yaxis As Vector
    Dim normal As Vector
    Dim newpos As Vector
    
    origin = Scalar(Axis, Dot(thevector, Axis) / Mag(Axis))
    xaxis = Subs(thevector, origin)
    normal = Scalar(Cross(thevector, Axis), 1 / (Mag(Axis) * Mag(thevector)))
    yaxis = Scalar(Scalar(normal, 1 / Mag(normal)), Mag(xaxis))
    newpos = Add(origin, Add(Scalar(xaxis, Cos(angle)), Scalar(yaxis, Sin(angle))))
    thevector.x = newpos.x
    thevector.y = newpos.y
    thevector.z = newpos.z
End Function

