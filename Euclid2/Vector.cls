VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Double
Public Y As Double

Public Sub Create(ByVal dX As Double, ByVal dY As Double)
    X = dX
    Y = dY
End Sub

Public Sub CreatePolar(ByVal dSize As Double, ByVal dAngle As Double)
    X = dSize * Cos(dAngle)
    Y = dSize * Sin(dAngle)
End Sub

Public Function Clone() As Vector
    Set Clone = New Vector
    Clone.X = X
    Clone.Y = Y
End Function

Public Function Copy(oVector)
    X = oVector.X
    Y = oVector.Y
End Function

Public Function Size() As Double
    Size = Sqr(X * X + Y * Y)
End Function

Public Function Dot(oVector As Vector) As Double
    Dot = X * oVector.X + Y * oVector.Y
End Function

Public Function Cross(oVector As Vector) As Double
    Cross = Y * oVector.X - X * oVector.Y
End Function

Public Function Normal() As Vector
    Dim dSize As Double
    
    Set Normal = New Vector
            
    dSize = Size
    If dSize <> 0 Then
        Normal.X = X / dSize
        Normal.Y = Y / dSize
    End If
End Function

Public Function Orthogonal() As Vector
    Dim dSize As Double
    
    dSize = Size
    Set Orthogonal = New Vector
    
    If dSize <> 0 Then
        Orthogonal.X = -Y / Size
        Orthogonal.Y = X / Size
    End If
End Function

Public Function Subtract(oVector As Vector) As Vector
    Set Subtract = New Vector
    
    Subtract.X = X - oVector.X
    Subtract.Y = Y - oVector.Y
End Function

Public Function Add(oVector) As Vector
    Set Add = New Vector
    
    Add.X = X + oVector.X
    Add.Y = Y + oVector.Y
End Function

Public Function Scaling(ByVal dScaling As Double) As Vector
    Set Scaling = New Vector

    Scaling.X = X * dScaling
    Scaling.Y = Y * dScaling
End Function

Public Function Distance(oVector As Vector) As Double
    Distance = Sqr((X - oVector.X) * (X - oVector.X) + (Y - oVector.Y) * (Y - oVector.Y))
End Function
