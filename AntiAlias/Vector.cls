VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Single
Public Y As Single

Public Function Create(dX As Long, dY As Long) As clsVector
    Set Create = New clsVector
    Create.X = dX
    Create.Y = dY
End Function

Public Function Add(Q As clsVector) As clsVector
    Set Add = New clsVector
    
    Add.X = X + Q.X
    Add.Y = Y + Q.Y
End Function

Public Function Subs(Q As clsVector) As clsVector
    Set Subs = New clsVector
    
    Subs.X = X - Q.X
    Subs.Y = Y - Q.Y
End Function

Public Function Scalar(ByVal S As Single) As clsVector
    Set Scalar = New clsVector
    
    Scalar.X = X * S
    Scalar.Y = Y * S
End Function

Public Property Get ArrayOf() As Variant
    ArrayOf = Array(X, Y)
End Property

Public Property Get Normal() As clsVector
    If Size() <> 0 Then
        Set Normal = Scalar(1 / Size())
    Else
        Set Normal = Scalar(0)
    End If
End Property

Public Function Dot(Q As clsVector) As Long
    Dot = X * Q.X + Y * Q.Y
End Function

Public Function Cross(Q As clsVector) As Long
    Cross = X * Q.Y - Y * Q.X
End Function

Public Function Size() As Single
    Size = Sqr(X * X + Y * Y)
End Function

Public Property Get Perpendicular() As clsVector
    Set Perpendicular = New clsVector
    Perpendicular.X = -Y
    Perpendicular.Y = X
End Property

Public Function Equal(a As clsVector) As Boolean
    If X = a.X And Y = a.Y Then
        Equal = True
    End If
End Function

Public Sub SetVector(ByVal dX As Long, ByVal dY As Long)
    X = dX
    Y = dY
End Sub

Public Sub Copy(Q As clsVector)
    X = Q.X
    Y = Q.Y
End Sub
