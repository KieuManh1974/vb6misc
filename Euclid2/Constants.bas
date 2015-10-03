Attribute VB_Name = "Constants"
Option Explicit

Public Const euColourNormal As Long = 0
Public Const euColourSelected As Long = vbRed

Public Enum ShapeStates
    ssNormal
    ssSelected
    ssDragged
    ssRemove
End Enum

Public Const TouchRadius As Double = 20

Public Enum LineStyles
    lsTrans = 0
    lsLeftDotted = 1
    lsLeftSolid = 2
    lsMidDotted = 3
    lsMidSolid = 6
    lsRightDotted = 9
    lsRightSolid = 18
End Enum
    
    
Public Const PI2 As Double = 6.28318530717959

Public Function Angle(oVector As Vector) As Double
    If oVector.X > 0 Then
        Angle = Atn(oVector.Y / oVector.X)
        If Angle < 0 Then
            Angle = Atn2 + PI2
        End If
    ElseIf oVector.X < 0 Then
        Angle = Atn(oVector.Y / oVector.X) - PI2 / 2
        If Angle < 0 Then
            Angle = Atn2 + PI2
        End If
    Else
        If oVector.Y > 0 Then
            Angle = PI2 / 4
        ElseIf oVector.Y < 0 Then
            Angle = 3 * PI2 / 4
        Else
            Angle = Rnd * PI2
        End If
    End If
End Function
