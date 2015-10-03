VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private oStaticVector As New Vector
Private fTimeStep As Double

Private Sub Form_Activate()
    Dim oPlanet1 As New Planet
    Dim oPlanet2 As New Planet
    Dim oPlanet3 As New Planet
    
    Dim oOffset As Vector
    Dim dTime As Double
    
    Dim dDateOffset  As Date
    
    dDateOffset = CDate("3 july 2006 23:00")
    
    Set oOffset = oStaticVector.Create(200, 200, 0)
    
    oPlanet1.Mass = 1.988435E+30 'Sun
    Set oPlanet1.Position = oStaticVector.Create(0, 0, 0)
    Set oPlanet1.Velocity = oStaticVector.Create(0, 0, 0)
        
    oPlanet2.Mass = 5.9742E+24 ' Earth
    Set oPlanet2.Position = oStaticVector.Create(152100000000#, 0, 0)
    Set oPlanet2.Velocity = oStaticVector.Create(0, -29290#, 0)
    oPlanet2.Tilt = pi2 * 23.439281 / 360
    oPlanet2.TiltOffset = pi2 * (6 * 15 + 11.8) / 360
    oPlanet2.Radius = 6372.795477598 * 1000
    oPlanet2.RateOfRevolution = pi2 / (23.9344696 * 60 * 60)
    oPlanet2.LongitudeOffset = -pi2 * (13 * 15 - 11.75) / 360
      
    oPlanet3.Mass = 6.4185E+23 ' Mars
    Set oPlanet3.Position = oStaticVector.Create(249230000000#, 0, 0)
    Set oPlanet3.Velocity = oStaticVector.Create(0, -21970#, 0)
        
    Me.ForeColor = vbWhite
    Do
        Me.Cls

        Debug.Print Format$(dDateOffset + dTime / 60 / 60 / 24, "DD/MM/YYYY HH:MM") & "   :   " & Format$(oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50 / 360, pi2 * 0 / 360, dTime).x, "0.000")
        'PlotPlanet oPlanet1.Position, oOffset, 1 / 1521000000, vbRed
        'PlotPlanet oPlanet2.Position, oOffset, 1 / 1521000000, vbWhite
        'PlotPlanet oPlanet3.Position, oOffset, 1 / 1521000000, vbWhite
        
        Set oPlanet1.Acceleration = oPlanet1.CurrentAcceleration(oPlanet2, oPlanet3)
        Set oPlanet2.Acceleration = oPlanet2.CurrentAcceleration(oPlanet1, oPlanet3)
        Set oPlanet3.Acceleration = oPlanet3.CurrentAcceleration(oPlanet1, oPlanet2)
        
        oPlanet1.UpdatePosition (fTimeStep)
        oPlanet2.UpdatePosition (fTimeStep)
        oPlanet3.UpdatePosition (fTimeStep)
        
        Me.Refresh
        DoEvents
        dTime = dTime + fTimeStep 'seconds
        If fTimeStep = 3600 Then
            fTimeStep = CDbl(12) * 3600
        Else
            fTimeStep = CDbl(12) * 3600
        End If
    Loop
End Sub

Private Sub PlotPlanet(oPosition As Vector, oOffset As Vector, dScaling As Double, lColour As Long)
    SetPixelV Me.hdc, CLng(oPosition.Scalar(dScaling).Add(oOffset).x), CLng(oPosition.Scalar(dScaling).Add(oOffset).y), lColour
End Sub

Private Sub Form_Load()
    pi2 = 8 * Atn(1)
    fTimeStep = 3600
End Sub

Private Function DMS(fValue As Double) As String
    DMS = Int(fValue) & " " & Int(fValue * 60) Mod 60 & " " & Int(fValue * 3600) Mod 60
End Function


