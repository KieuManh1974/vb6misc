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

Public x As Double
Public y As Double
Public z As Double

Public Sub Vector(dX As Double, dY As Double, dZ As Double)
    x = dX
    y = dY
    z = dZ
End Sub

Public Function Create(dX As Double, dY As Double, dZ As Double) As Vector
    Set Create = New Vector
    
    Create.x = dX
    Create.y = dY
    Create.z = dZ
End Function

Public Function Add(Q As Vector) As Vector
    Set Add = New Vector
    
    Add.x = x + Q.x
    Add.y = y + Q.y
    Add.z = z + Q.z
End Function

Public Function Subs(Q As Vector) As Vector
    Set Subs = New Vector
    
    Subs.x = x - Q.x
    Subs.y = y - Q.y
    Subs.z = z - Q.z
End Function

Public Function Scalar(S As Double) As Vector
    Set Scalar = New Vector
    
    Scalar.x = x * S
    Scalar.y = y * S
    Scalar.z = z * S
End Function

Public Property Get ArrayOf() As Variant
    ArrayOf = Array(x, y, z)
End Property

Public Property Get Normal() As Vector
    Set Normal = Scalar(1 / Size())
End Property

Public Function Cross(Q As Vector) As Vector
    Set Cross = New Vector
    Cross.x = y * Q.z - z * Q.y
    Cross.y = z * Q.x - x * Q.z
    Cross.z = x * Q.y - y * Q.x
    
    Set Cross = Cross.Scalar(1 / (Size * Q.Size))
End Function

Public Function Size() As Double
    Size = Sqr(x * x + y * y + z * z)
End Function

Public Property Get Flat(iIndex As Integer) As Vector
    Set Flat = New Vector
    
    Flat.x = (x / (1000 - z)) * 10000 + 5000
    Flat.y = (-y / (1000 - z)) * 10000 + 5000
End Property

Public Function Multiply(Q As Vector) As Vector
    Set Multiply = New Vector
    
    Multiply.x = x * Q.x
    Multiply.y = y * Q.y
    Multiply.z = z * Q.z
End Function

Public Function Dot(Q As Vector) As Double
    Dot = x * Q.x + y * Q.y + z * Q.z
End Function

Public Function RotateZ(A As Double) As Vector
    Set RotateZ = New Vector
    
    RotateZ.x = Cos(A) * x - Sin(A) * y
    RotateZ.y = Sin(A) * x + Cos(A) * y
    RotateZ.z = z
End Function

Public Function RotateX(A As Double) As Vector
    Set RotateX = New Vector
    
    RotateX.x = x
    RotateX.y = Sin(A) * z + Cos(A) * y
    RotateX.z = Cos(A) * z - Sin(A) * y
End Function

Public Function RotateY(A As Double) As Vector
    Set RotateY = New Vector
    
    RotateY.x = Cos(A) * x - Sin(A) * z
    RotateY.y = y
    RotateY.z = Sin(A) * x + Cos(A) * z
End Function

Public Function RotateAbout(Q As Vector, A As Double) As Vector
    Dim fCosY As Double
    Dim fCosZ As Double
    Dim fSinY As Double
    Dim fSinZ As Double
    Dim vRotateZ As Vector
    
    fCosY = Q.x / Create(Q.x, Q.y, 0).Size
    fSinY = -Q.y / Create(Q.x, Q.y, 0).Size
    
    Set vRotateZ = Create(fCosY * Q.x - fSinY * Q.y, fSinY * Q.x + fCosY * Q.y, Q.z)
    
    fCosZ = vRotateZ.x / Create(vRotateZ.x, 0, vRotateZ.z)
    fSinZ = -vRotateZ.y / Create(vRotateZ.x, 0, vRotateZ.z)
    
    
End Function

Public Function Description() As String
    Description = "[" & Format$(x, "0.000") & " , " & Format$(y, "0.000") & " , " & Format$(z, "0.000") & "]"
End Function
