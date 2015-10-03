VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Private mlDX As Long
Private mlDY As Long
Private mlSubDX As Long
Private mlSubDY As Long
Private mlScanlineDisc As Long
Private mlPixelDisc As Long
Private mlSubScanlineDisc As Long
Private mlSubPixelDisc As Long
Private mnRadius As Single
Private moCentre As clsVector
Private mlColour As Long
Private mbInverse As Boolean

Private mlRed As Long
Private mlGreen As Long
Private mlBlue As Long

Public Property Get Centre() As clsVector
    Set Centre = moCentre
End Property

Public Property Get Radius() As Single
    Radius = mnRadius
End Property

Public Sub SetUp(oCentre As clsVector, ByVal nRadius As Single)
    Set moCentre = oCentre.Scalar(16)
    mnRadius = nRadius * 16
End Sub


Public Property Let IShape_Colour(ByVal lColour As Long)
    mlColour = lColour
    
    mlRed = lColour And &HFF
    mlGreen = (lColour \ 256) And &HFF
    mlBlue = (lColour \ 65536) And &HFF
End Property

Public Property Get IShape_Colour() As Long
    IShape_Colour = mlColour
End Property

Private Property Get IShape_Red() As Long
    IShape_Red = mlRed
End Property

Private Property Get IShape_Green() As Long
    IShape_Green = mlGreen
End Property

Private Property Get IShape_Blue() As Long
    IShape_Blue = mlBlue
End Property


Public Property Let IShape_Inverse(ByVal bInverse As Boolean)
    mbInverse = bInverse
End Property


Public Sub IShape_StartScan()
    mlDX = -moCentre.X
    mlDY = -moCentre.Y
    mlScanlineDisc = mlDX * mlDX + mlDY * mlDY - mnRadius * mnRadius
    mlPixelDisc = mlScanlineDisc
    
    mlSubDX = mlDX
    mlSubDY = mlDY
    mlSubScanlineDisc = mlScanlineDisc
    mlSubPixelDisc = mlPixelDisc
End Sub

Public Function IShape_Inside() As Boolean
    If mbInverse Then
        IShape_Inside = (mlSubPixelDisc > 0)
    Else
        IShape_Inside = (mlSubPixelDisc < 0)
    End If
End Function

Private Sub IShape_NextSubPixel()
    mlSubPixelDisc = mlSubPixelDisc + 2 * mlSubDX + 1
    mlSubDX = mlSubDX + 1
End Sub

Private Sub IShape_NextSubScanline()
    mlSubScanlineDisc = mlSubScanlineDisc + 2 * mlSubDY + 1
    mlSubPixelDisc = mlSubScanlineDisc
    mlSubDY = mlSubDY + 1
    mlSubDX = mlDX
End Sub

Public Sub IShape_NextScanline()
    mlScanlineDisc = mlScanlineDisc + 32 * mlDY + 256
    mlPixelDisc = mlScanlineDisc
    mlDY = mlDY + 16
    mlDX = -moCentre.X
    
    mlSubScanlineDisc = mlScanlineDisc
    mlSubPixelDisc = mlPixelDisc
    mlSubDY = mlDY
    mlSubDX = mlDX
End Sub

Public Sub IShape_NextPixel()
    mlPixelDisc = mlPixelDisc + 32 * mlDX + 256
    mlDX = mlDX + 16
    
    mlSubScanlineDisc = mlPixelDisc
    mlSubPixelDisc = mlPixelDisc
    mlSubDX = mlDX
    mlSubDY = mlDY
End Sub

