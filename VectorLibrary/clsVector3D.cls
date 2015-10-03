VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Explicit

Public X As Double
Public Y As Double
Public Z As Double

Public Function Create(dX As Double, dY As Double, dZ As Double) As clsVector3d
    Set Create = New clsVector3d
    
    Create.X = dX
    Create.Y = dY
    Create.Z = dZ
End Function

Public Function Add(Q As clsVector3d) As clsVector3d
    Set Add = New clsVector3d
    
    Add.X = X + Q.X
    Add.Y = Y + Q.Y
    Add.Z = Z + Q.Z
End Function

Public Function Subs(Q As clsVector3d) As clsVector3d
    Set Subs = New clsVector3d
    
    Subs.X = X - Q.X
    Subs.Y = Y - Q.Y
    Subs.Z = Z - Q.Z
End Function

Public Function Scalar(S As Double) As clsVector3d
    Set Scalar = New clsVector3d
    
    Scalar.X = X * S
    Scalar.Y = Y * S
    Scalar.Z = Z * S
End Function

Public Property Get ArrayOf() As Variant
    ArrayOf = Array(X, Y, Z)
End Property

Public Property Get Normal() As clsVector3d
    Set Normal = Scalar(1 / Size())
End Property

Public Function Cross(Q As clsVector3d) As clsVector3d
    Set Cross = New clsVector3d
    Cross.X = Y * Q.Z - Z * Q.Y
    Cross.Y = Z * Q.X - X * Q.Z
    Cross.Z = X * Q.Y - Y * Q.X
    
    Set Cross = Cross.Scalar(1 / (Size * Q.Size))
End Function

Public Function Size() As Double
    Size = Sqr(X * X + Y * Y + Z * Z)
End Function

Public Property Get Flat(iIndex As Integer) As clsVector3d
    Set Flat = New clsVector3d
    
    Flat.X = (X / (1000 - Z)) * 10000 + 5000
    Flat.Y = (-Y / (1000 - Z)) * 10000 + 5000
End Property

Public Function Multiply(Q As clsVector3d) As clsVector3d
    Set Multiply = New clsVector3d
    
    Multiply.X = X * Q.X
    Multiply.Y = Y * Q.Y
    Multiply.Z = Z * Q.Z
End Function

Public Function Dot(Q As clsVector3d) As Double
    Dot = X * Q.X + Y * Q.Y + Z * Q.Z
End Function

Public Function RotateZ(A As Double) As clsVector3d
    Set RotateZ = New clsVector3d
    
    RotateZ.X = Cos(A) * X - Sin(A) * Y
    RotateZ.Y = Sin(A) * X + Cos(A) * Y
    RotateZ.Z = Z
End Function

Public Function RotateX(A As Double) As clsVector3d
    Set RotateX = New clsVector3d
    
    RotateX.X = X
    RotateX.Y = Sin(A) * Z + Cos(A) * Y
    RotateX.Z = Cos(A) * Z - Sin(A) * Y
End Function

Public Function RotateY(A As Double) As clsVector3d
    Set RotateY = New clsVector3d
    
    RotateY.X = Cos(A) * X - Sin(A) * Z
    RotateY.Y = Y
    RotateY.Z = Sin(A) * X + Cos(A) * Z
End Function

Public Function RotateAbout(Q As clsVector3d, A As Double) As clsVector3d
    Dim fCosY As Double
    Dim fCosZ As Double
    Dim fSinY As Double
    Dim fSinZ As Double
    Dim vRotateZ As clsVector3d
    
    fCosY = Q.X / Create(Q.X, Q.Y, 0).Size
    fSinY = -Q.Y / Create(Q.X, Q.Y, 0).Size
    
    Set vRotateZ = Create(fCosY * Q.X - fSinY * Q.Y, fSinY * Q.X + fCosY * Q.Y, Q.Z)
    
    fCosZ = vRotateZ.X / Create(vRotateZ.X, 0, vRotateZ.Z)
    fSinZ = -vRotateZ.Y / Create(vRotateZ.X, 0, vRotateZ.Z)
    
    
End Function

Public Function Description() As String
    Description = "[" & Format$(X, "0.000") & " , " & Format$(Y, "0.000") & " , " & Format$(Z, "0.000") & "]"
End Function

