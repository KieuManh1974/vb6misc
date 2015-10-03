VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Double
Public Y As Double

Public Function Create(dX As Double, dY As Double) As clsVector2D
    Set Create = New clsVector2D
    X = dX
    Y = dY
End Function

Public Function Add(Q As Vector) As clsVector2D
    Set Add = New clsVector2D
    
    Add.X = X + Q.X
    Add.Y = Y + Q.Y
End Function

Public Function Subs(Q As Vector) As clsVector2D
    Set Subs = New clsVector2D
    
    Subs.X = X - Q.X
    Subs.Y = Y - Q.Y
End Function

Public Function Scalar(ByVal S As Double) As clsVector2D
    Set Scalar = New clsVector2D
    
    Scalar.X = X * S
    Scalar.Y = Y * S
End Function

Public Property Get ArrayOf() As Variant
    ArrayOf = Array(X, Y)
End Property

Public Property Get Normal() As clsVector2D
    Set Normal = Scalar(1 / Size())
End Property

Public Function Cross(Q As Vector) As clsVector2D
    Set Cross = New clsVector2D
    
    Cross.X = Y * Q.X - X * Q.Y
    Cross.Y = X * Q.Y - Y * Q.X
    
    Set Cross = Cross.Scalar(1 / (Size * Q.Size))
End Function

Public Function Size() As Double
    Size = Sqr(X * X + Y * Y)
End Function

Public Function Multiply(Q As clsVector2D) As clsVector2D
    Set Multiply = New clsVector2D
    
    Multiply.X = X * Q.X
    Multiply.Y = Y * Q.Y
End Function

Public Function Dot(Q As clsVector3d) As Double
    Dot = X * Q.X + Y * Q.Y
End Function

Public Function Rotate(A As Double) As clsVector2D
    Set Rotate = New clsVector3d
    
    Rotate.X = Cos(A) * X - Sin(A) * Y
    Rotate.Y = Sin(A) * X + Cos(A) * Y
End Function

Public Property Get Perpendicular() As clsVector2D
    Set Perpendicular = New Vector
    Perpendicular.X = -Y
    Perpendicular.Y = X
End Property

Public Function Equal(A As clsVector2D) As Boolean
    If X = A.X And Y = A.Y Then
        Equal = True
    End If
End Function

Public Function Description() As String
    Description = "[" & Format$(X, "0.000") & " , " & Format$(Y, "0.000") & "]"
End Function
