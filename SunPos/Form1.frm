VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   14850
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11505
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub SetTextAlign Lib "gdi32" (ByValhDC As Long, ByValwFlags As Long)

Private oStaticVector As New Vector

Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const TA_CENTER = 6
Private Const TA_TOP = 0
Private Const TA_BOTTOM = 8
Private Const TA_BASELINE = 24

Sub main()
    Debug.Print SunPosition("21 dec 2006 12:00:00").x
End Sub

Private Function DMS(fValue As Double) As String
    DMS = Int(fValue) & " " & Int(fValue * 60) Mod 60 & " " & Int(fValue * 3600) Mod 60
End Function

Public Sub PlotSunPositions()
    Dim oPlanet1 As New Planet
    Dim oPlanet2 As New Planet
    
    Dim fTime As Double
    Dim fStep As Double
    
    Dim dPerihelion  As Date
    Dim oPolarPosition As Vector
    Dim fShadowLength As Double
    Dim fScaling As Double
    
    Dim lMinute As Long
    Dim lDay As Long
    
    Dim nOffsetX As Single
    Dim nOffsetY As Single
    
    Dim oPos(1439, 366) As Vector
    
    dPerihelion = CDate("3 july 2006 23:00")
    
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
    
    fStep = 60
    fScaling = 1200
'    fScaling = 100
    Printer.ScaleMode = vbPixels
    lMinute = 23 * 60
    lDay = 0
    
    nOffsetX = Printer.Width / Printer.TwipsPerPixelX / 2 - 2000
    nOffsetY = Printer.Height / Printer.TwipsPerPixelY / 2
'    nOffsetX = Me.Width / Screen.TwipsPerPixelX / 2 - 100
'    nOffsetY = Me.Height / Screen.TwipsPerPixelY / 2
    
    For fTime = 0 To CDbl(3600) * 24 * 366 Step fStep
        Set oPolarPosition = oPlanet2.PolarPosition(oPlanet1.Position, pi2 * 50.805 / 360, pi2 * -0.034 / 360, fTime)
        'Debug.Print Format$(dPerihelion + fTime / 60 / 60 / 24, "DD/MM/YYYY HH:MM:SS") & "   :   " & Format$(oPolarPosition.x, "0.000") & "   :   " & Format$(oPolarPosition.y, "0.000")

        If oPolarPosition.x >= 0 Then
            fShadowLength = fScaling / Tan(oPolarPosition.x)
            
            Set oPos(lMinute, lDay) = New Vector
            oPos(lMinute, lDay).x = nOffsetX - Cos(oPolarPosition.y) * fShadowLength
            oPos(lMinute, lDay).y = nOffsetY - Sin(oPolarPosition.y) * fShadowLength
        End If
        
        Set oPlanet2.Acceleration = oPlanet2.CurrentAcceleration(oPlanet1)
        oPlanet2.UpdatePosition (fStep)
        lMinute = (lMinute + 1) Mod 1440
        If lMinute = 0 Then
            lDay = lDay + 1
            Debug.Print lDay
        End If
    Next

    On Error Resume Next
    For lMinute = 0 To 1438 Step 1
        For lDay = 0 To 366
            If lDay = 1 Then
                If lMinute Mod 60 = 0 Then
                    If Not oPos(lMinute, lDay) Is Nothing Then
                        Printer.CurrentX = oPos(lMinute, lDay).x - Me.TextWidth(lMinute \ 60) * 2
                        Printer.CurrentY = oPos(lMinute, lDay).y - Me.TextHeight("X") / 2
                        Printer.Print lMinute \ 60
                        
'                        Me.CurrentX = oPos(lMinute, lDay).x - Me.TextWidth(lMinute \ 60) * 2
'                        Me.CurrentY = oPos(lMinute, lDay).y - Me.TextHeight("X") / 2
'                        Me.Print lMinute \ 60
                    End If
                End If
            End If
            If lMinute Mod 30 = 0 Then
                If Not oPos(lMinute, lDay) Is Nothing And Not oPos(lMinute, lDay + 1) Is Nothing Then
                    Printer.Line (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y)-(oPos(lMinute, lDay + 1).x, oPos(lMinute, lDay + 1).y), vbBlack
'                    Me.Line (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y)-(oPos(lMinute, lDay + 1).x, oPos(lMinute, lDay + 1).y), vbBlack
                End If
            ElseIf Not oPos(lMinute, lDay) Is Nothing Then
                Printer.PSet (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y)
                Printer.PSet (oPos(lMinute, lDay).x + 1, oPos(lMinute, lDay).y)
                Printer.PSet (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y + 1)
                Printer.PSet (oPos(lMinute, lDay).x + 1, oPos(lMinute, lDay).y + 1)
            End If
        Next
    Next
    
    For lDay = 333 To 333
        For lMinute = 0 To 1438
            If Not oPos(lMinute, lDay) Is Nothing And Not oPos(lMinute + 1, lDay) Is Nothing Then
                Printer.Line (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y)-(oPos(lMinute + 1, lDay).x, oPos(lMinute + 1, lDay).y), vbBlack
'                Me.Line (oPos(lMinute, lDay).x, oPos(lMinute, lDay).y)-(oPos(lMinute + 1, lDay).x, oPos(lMinute + 1, lDay).y), vbBlack
            End If
        Next
    Next

    Printer.Line (nOffsetX - fScaling, nOffsetY)-Step(50 + fScaling, 0), vbBlack
    Printer.Line (nOffsetX, nOffsetY - 50)-Step(0, 100), vbBlack
'    Me.Line (nOffsetX - fScaling, nOffsetY)-Step(100 / Screen.TwipsPerPixelX + fScaling, 0), vbBlack
'    Me.Line (nOffsetX, nOffsetY - 100 / Screen.TwipsPerPixelX)-Step(0, 200 / Screen.TwipsPerPixelX), vbBlack
    
    Printer.EndDoc
End Sub

Private Sub Form_Activate()
    SetTextAlign Me.hDC, TA_RIGHT Or TA_CENTER
    PlotSunPositions
End Sub

Private Sub Form_Load()
    pi2 = 8 * Atn(1)
End Sub
