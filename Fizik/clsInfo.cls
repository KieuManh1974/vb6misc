VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moPosition As New clsVector
Private moVelocity As New clsVector
Private moForce As New clsVector
Private moAcceleration As New clsVector
Private mfMass As Double

Public Property Get Position() As clsVector
    Set Position = moPosition
End Property

Public Property Get Velocity() As clsVector
    Set Velocity = moVelocity
End Property

Public Property Get Force() As clsVector
    Set Force = moForce
End Property

Public Property Get Acceleration() As clsVector
    Set Acceleration = moAcceleration
End Property

Public Property Get Mass() As Double
    Mass = mfMass
End Property

Public Property Set Position(ByVal oPosition As clsVector)
    Set moPosition = oPosition
End Property

Public Property Set Velocity(ByVal oVelocity As clsVector)
    Set moVelocity = oVelocity
End Property

Public Property Set Force(ByVal oForce As clsVector)
    Set moForce = oForce
End Property

Public Property Set Acceleration(ByVal oAcceleration As clsVector)
    Set moAcceleration = oAcceleration
End Property

Public Property Let Mass(ByVal fMass As Double)
    mfMass = fMass
End Property

Public Sub UpdateMotion(ByVal fInterval As Double)
    Dim lIndex As Long
    
    For lIndex = 0 To 1
        moAcceleration.Component(lIndex) = moForce.Component(lIndex) / mfMass
        moVelocity.Component(lIndex) = moVelocity.Component(lIndex) + moAcceleration.Component(lIndex)
        moPosition.Component(lIndex) = moPosition.Component(lIndex) + moVelocity.Component(lIndex)
    Next
End Sub
