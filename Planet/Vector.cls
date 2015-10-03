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
Public x As Double
Public y As Double

Public Property Get size() As Double
    size = Sqr(x * x + y * y)
End Property

Public Property Get angle() As Double
    quadrant = (Sgn(x) + 1) \ 2 + (Sgn(y) + 1)
        
    If x = 0 Then
        If y > 0 Then
        End If
    End If
    If y = 0 Then
        If x > 0 Then
        End If
    End If
    
    
End Property

Public Function subtract_self(v2 As Vector) As Vector
    x = x - v2.x
    y = y - v2.y
End Function

Public Function add_self(v2 As Vector) As Vector
    x = x + v2.x
    y = y + v2.y
End Function


Public Function addition(v2 As Vector) As Vector
    addition = New Vector
    addition.x = x + v2.y
    addition.y = y + v2.x
End Function


Public Function subtract(v2 As Vector) As Vector
    subtract = New Vector
    subtract.x = x - v2.y
    subtract.y = y - v2.x
End Function
