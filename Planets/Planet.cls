VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Planet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Mass As Double
Public Position As New Vector
Public Velocity As New Vector
Public Acceleration As New Vector

Public Tilt As Double
Public TiltOffset As Double
Public RateOfRevolution As Double
Public Radius As Double
Public LongitudeOffset As Double

Public Sub UpdatePosition(dInterval As Double)
    Set Velocity = Velocity.Add(Acceleration.Scalar(dInterval))
    Set Position = Position.Add(Velocity.Scalar(dInterval))
End Sub

Public Function CurrentAcceleration(ParamArray oPlanets() As Variant) As Vector
    Dim oPlanet As Variant
    
    Set CurrentAcceleration = New Vector
    For Each oPlanet In oPlanets
        Set CurrentAcceleration = CurrentAcceleration.Add(oPlanet.Position.Subs(Position).Normal.Scalar(UGA * oPlanet.Mass / (oPlanet.Position.Subs(Position).Size ^ 2)))
    Next
End Function

Public Function SurfacePosition(ByVal fLatitude As Double, ByVal fLongitude As Double, ByVal fTime As Double) As Vector
    Set SurfacePosition = New Vector
    
    SurfacePosition.x = Sin(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Cos(fLatitude)
    SurfacePosition.y = Cos(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Cos(fLatitude)
    SurfacePosition.z = Sin(fLatitude)

    Set SurfacePosition = SurfacePosition.Scalar(Radius).RotateX(Tilt).RotateZ(TiltOffset)
End Function


Public Function PolarPosition(oBody As Vector, ByVal fLatitude As Double, ByVal fLongitude As Double, ByVal fTime As Double) As Vector
    Dim oV1 As Vector
    Dim oV2 As Vector
    Dim fAltitude As Double
    Dim fAzimuth As Double
    Dim oSP As Vector
    
    Set oSP = SurfacePosition(fLatitude, fLongitude, fTime)
    Set oV1 = oBody.Subs(Position).Subs(oSP)

    ' Dot product to find angular height
    fAltitude = ACS(oSP.Dot(oV1) / oSP.Size / oV1.Size)
    
    Set PolarPosition = New Vector
    Set PolarPosition = PolarPosition.Create(90 - 360 * fAltitude / pi2, fAzimuth, 0)
    
End Function

