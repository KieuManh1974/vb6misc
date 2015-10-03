Attribute VB_Name = "Calculate"
Option Explicit

Private oStaticVector As New Vector

Sub main()
    Debug.Print SunPosition("21 dec 2006 12:00:00").x
End Sub

Private Function DMS(fValue As Double) As String
    DMS = Int(fValue) & " " & Int(fValue * 60) Mod 60 & " " & Int(fValue * 3600) Mod 60
End Function

Private Function SunPosition(ByVal dDate As Date) As Variant
    Dim oPlanet1 As New Planet
    Dim oPlanet2 As New Planet
    
    Dim oOffset As Vector
    Dim dTime As Double
    Dim fTimeStep As Double

    Dim dPerihelion  As Date
    Dim oPolarPosition As Vector
    
    dPerihelion = CDate("3 july 2006 23:00")
    

    Set oOffset = oStaticVector.Create(200, 200, 0)
    
    oPlanet1.Mass = 1.988435E+30 'Sun
    Set oPlanet1.Position = oStaticVector.Create(0, 0, 0)
    Set oPlanet1.Velocity = oStaticVector.Create(0, 0, 0)
        
    oPlanet2.Mass = 5.9742E+24 ' Earth
    Set oPlanet2.Position = oStaticVector.Create(152100000000#, 0, 0)
    Set oPlanet2.Velocity = oStaticVector.Create(0, 29290#, 0)
    oPlanet2.Tilt = pi2 * 23.439281 / 360
    oPlanet2.TiltOffset = pi2 * (-11.523) / 360
    oPlanet2.Radius = 6372.795477598 * 1000
    oPlanet2.RateOfRevolution = pi2 / (23.9344696 * 60 * 60)
    oPlanet2.LongitudeOffset = pi2 * (-12 * 15 - 3) / 360
    
    fTimeStep = 60 ' seconds
    
    
    Do
        Set oPolarPosition = oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime)
        'Debug.Print Format$(dPerihelion + dTime / 60 / 60 / 24, "DD/MM/YYYY HH:MM:SS") & "   :   " & Format$(oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime).x, "0.000") & "   :   " & Format$(oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime).y, "0.000")

        Set oPlanet2.Acceleration = oPlanet2.CurrentAcceleration(oPlanet1)
        
'        If ((dTime + fTimeStep) / 86400 + dPerihelion) > dDate Then
'            fTimeStep = fTimeStep / 60
'        End If

        oPlanet2.UpdatePosition (fTimeStep)
        
        dTime = dTime + fTimeStep 'seconds
        
'        If fTimeStep = 3600 Then
'            fTimeStep = CDbl(11) * 3600 + 55 * 60
'        ElseIf fTimeStep = CDbl(11) * 3600 + 55 * 60 Then
'            fTimeStep = 60
'        End If
    Loop Until Abs((dTime / 86400 + dPerihelion) - dDate) < 1 / 86400
    
    Set SunPosition = oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime)
    Debug.Print Format$(dPerihelion + dTime / 60 / 60 / 24, "DD/MM/YYYY HH:MM:SS") & "   :   " & Format$(oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime).x, "0.000") & "   :   " & Format$(oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime).y, "0.000")
    'SunPosition = Format$(oSunPos.x, "0.000") & "   :   " & Format$(oSunPos.y, "0.000")
End Function

Private Sub Form_Load()
    pi2 = 8 * Atn(1)
End Sub

