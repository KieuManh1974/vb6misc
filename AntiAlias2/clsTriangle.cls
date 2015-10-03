VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Private moPointA As clsVector
Private moPointB As clsVector
Private moPointC As clsVector

Private moLine1 As clsVector
Private moLine2 As clsVector
Private moLine3 As clsVector

Private mlScanlineDisc1 As Long
Private mlScanlineDisc2 As Long
Private mlScanlineDisc3 As Long

Private mlPixelDisc1 As Long
Private mlPixelDisc2 As Long
Private mlPixelDisc3 As Long

Private mlSubScanlineDisc1 As Long
Private mlSubScanlineDisc2 As Long
Private mlSubScanlineDisc3 As Long

Private mlSubPixelDisc1 As Long
Private mlSubPixelDisc2 As Long
Private mlSubPixelDisc3 As Long

Private mlSign1 As Long
Private mlSign2 As Long
Private mlSign3 As Long

Private mlColour As Long
Private mlRed As Long
Private mlGreen As Long
Private mlBlue As Long

Private mlInverse As Long

Private Sub Class_Initialize()
    mlInverse = 1
End Sub

Public Property Get Point(ByVal lIndex As Long) As clsVector
    Select Case lIndex
        Case 1
            Set Point = moPointA
        Case 2
            Set Point = moPointB
        Case 3
            Set Point = moPointC
    End Select
End Property

Public Property Let IShape_Colour(ByVal lColour As Long)
    mlColour = lColour
    
    mlRed = lColour And &HFF
    mlGreen = (lColour \ 256) And &HFF
    mlBlue = (lColour \ 65536) And &HFF
End Property

Public Property Get IShape_Colour() As Long
    IShape_Colour = mlColour
End Property


Public Property Get IShape_Red() As Long
    IShape_Red = mlRed
End Property

Public Property Get IShape_Green() As Long
    IShape_Green = mlGreen
End Property

Public Property Get IShape_Blue() As Long
    IShape_Blue = mlBlue
End Property


Public Property Let IShape_Inverse(ByVal bInverse As Boolean)
    mlInverse = bInverse * 2 + 1
End Property

Public Sub SetUp(oPointA As clsVector, oPointB As clsVector, oPointC As clsVector)
    Set moPointA = oPointA.Scalar(16)
    Set moPointB = oPointB.Scalar(16)
    Set moPointC = oPointC.Scalar(16)
    
    Set moLine1 = oPointB.Subs(oPointA)
    Set moLine2 = oPointC.Subs(oPointB)
    Set moLine3 = oPointA.Subs(oPointC)
    
    mlSign1 = -Sgn(moLine3.Cross(moLine1))
    mlSign2 = -Sgn(moLine1.Cross(moLine2))
    mlSign3 = -Sgn(moLine2.Cross(moLine3))
End Sub


Public Sub IShape_StartScan()
    Dim oOrigin As New clsVector
    
    oOrigin.SetVector 0, 0
    
    mlScanlineDisc1 = oOrigin.Subs(moPointA).Cross(moLine1)
    mlPixelDisc1 = mlScanlineDisc1
    mlScanlineDisc2 = oOrigin.Subs(moPointB).Cross(moLine2)
    mlPixelDisc2 = mlScanlineDisc2
    mlScanlineDisc3 = oOrigin.Subs(moPointC).Cross(moLine3)
    mlPixelDisc3 = mlScanlineDisc3
End Sub

Public Sub IShape_NextScanline()
    mlScanlineDisc1 = mlScanlineDisc1 - moLine1.X * 16
    mlPixelDisc1 = mlScanlineDisc1
    mlScanlineDisc2 = mlScanlineDisc2 - moLine2.X * 16
    mlPixelDisc2 = mlScanlineDisc2
    mlScanlineDisc3 = mlScanlineDisc3 - moLine3.X * 16
    mlPixelDisc3 = mlScanlineDisc3
    
    mlSubScanlineDisc1 = mlScanlineDisc1
    mlSubScanlineDisc2 = mlScanlineDisc2
    mlSubScanlineDisc3 = mlScanlineDisc3
    
    mlSubPixelDisc1 = mlScanlineDisc1
    mlSubPixelDisc2 = mlScanlineDisc2
    mlSubPixelDisc3 = mlScanlineDisc3
End Sub

Public Sub IShape_NextPixel()
    mlPixelDisc1 = mlPixelDisc1 + moLine1.Y * 16
    mlPixelDisc2 = mlPixelDisc2 + moLine2.Y * 16
    mlPixelDisc3 = mlPixelDisc3 + moLine3.Y * 16
    
    mlSubScanlineDisc1 = mlPixelDisc1
    mlSubScanlineDisc2 = mlPixelDisc2
    mlSubScanlineDisc3 = mlPixelDisc3
    
    mlSubPixelDisc1 = mlPixelDisc1
    mlSubPixelDisc2 = mlPixelDisc2
    mlSubPixelDisc3 = mlPixelDisc3
End Sub

Private Sub IShape_NextSubPixel()
    mlSubPixelDisc1 = mlSubPixelDisc1 + moLine1.Y
    mlSubPixelDisc2 = mlSubPixelDisc2 + moLine2.Y
    mlSubPixelDisc3 = mlSubPixelDisc3 + moLine3.Y
End Sub

Private Sub IShape_NextSubScanline()
    mlSubScanlineDisc1 = mlSubScanlineDisc1 - moLine1.X
    mlSubPixelDisc1 = mlSubScanlineDisc1
    mlSubScanlineDisc2 = mlSubScanlineDisc2 - moLine2.X
    mlSubPixelDisc2 = mlSubScanlineDisc2
    mlSubScanlineDisc3 = mlSubScanlineDisc3 - moLine3.X
    mlSubPixelDisc3 = mlSubScanlineDisc3
End Sub

Public Function IShape_Inside() As Boolean
    IShape_Inside = ((mlSubPixelDisc1 * mlSign1 > 0) And (mlSubPixelDisc2 * mlSign2 > 0) And (mlSubPixelDisc3 * mlSign3 > 0)) * mlInverse
End Function

