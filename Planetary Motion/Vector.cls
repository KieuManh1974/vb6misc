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

Public Function Create(ByVal fX As Double, ByVal fY As Double)
    X = fX
    Y = fY
End Function

Public Function Cross(ByVal oVector As Vector) As Double
    Cross = X * oVector.Y - Y * oVector.X
End Function

Public Function Dot(ByVal oVector As Vector) As Double
    Dot = X * oVector.X + Y * oVector.Y
End Function

Public Function Size() As Double
    Size = Sqr(X * X + Y * Y)
End Function
