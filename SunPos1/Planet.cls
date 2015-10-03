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
    
    SurfacePosition.x = -Cos(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Cos(fLatitude)
    SurfacePosition.y = -Sin(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Cos(fLatitude)
    SurfacePosition.z = Sin(fLatitude)

    Set SurfacePosition = SurfacePosition.Scalar(Radius).RotateY(Tilt).RotateZ(TiltOffset)
End Function

Public Function NorthVector(ByVal fLatitude As Double, ByVal fLongitude As Double, ByVal fTime As Double) As Vector
    Set NorthVector = New Vector
    
    NorthVector.x = Cos(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Sin(fLatitude)
    NorthVector.y = Sin(fLongitude + LongitudeOffset + fTime * RateOfRevolution) * Sin(fLatitude)
    NorthVector.z = Cos(fLatitude)

    Set NorthVector = NorthVector.RotateY(Tilt).RotateZ(TiltOffset)
End Function

Public Function PolarPosition(oBody As Vector, ByVal fLatitude As Double, ByVal fLongitude As Double, ByVal fTime As Double) As Vector
    Dim oBodyVector As Vector
    Dim fAltitude As Double
    Dim fAzimuth As Double
    Dim oSurfaceNormal As Vector
    Dim oNorthVector As Vector
    Dim oEastVector As Vector
    Dim oProtoAzimuth As Vector
    Dim fAzimuth1 As Double
    Dim fAzimuth2 As Double
    
    Set oSurfaceNormal = SurfacePosition(fLatitude, fLongitude, fTime)
    Set oNorthVector = NorthVector(fLatitude, fLongitude, fTime)
    Set oEastVector = oNorthVector.Cross(oSurfaceNormal).Normal
    Set oBodyVector = oBody.Subs(Position).Subs(oSurfaceNormal)
    
    ' Dot product to find angular height
    fAltitude = ACS(oSurfaceNormal.Dot(oBodyVector) / oSurfaceNormal.Size / oBodyVector.Size)
    
    ' Cross product to find vector perpendicular to surface normal and sun position (this will lie on the plane)
    Set oProtoAzimuth = oSurfaceNormal.Cross(oBodyVector).Normal
        
    fAzimuth1 = ACS(oNorthVector.Dot(oProtoAzimuth))
    fAzimuth2 = oEastVector.Dot(oProtoAzimuth)
    
    If fAzimuth2 < 0 Then
        fAzimuth1 = pi2 - fAzimuth1
    End If
    
    Set PolarPosition = New Vector
    Set PolarPosition = PolarPosition.Create(pi2 / 4 - fAltitude, Modus(pi2 / 4 + fAzimuth1, pi2), 0)
End Function

Private Function Modus(ByVal fValue As Double, ByVal fRange As Double) As Double
    Dim fMultiplier As Double
    
    fMultiplier = Int(fValue / fRange)
    Modus = fValue - fRange * fMultiplier
End Function
