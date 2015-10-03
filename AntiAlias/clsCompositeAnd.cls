VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompositeAnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Private moShapes As New Collection

Private mlInverse As Long
Private mlColour As Long
Private mlRed As Long
Private mlGreen As Long
Private mlBlue As Long
Private mbInverse As Boolean

Public Function AddShape(oShape As IShape)
    moShapes.Add oShape
End Function

Public Property Get IShape_Colour() As Long
    IShape_Colour = mlColour
    
End Property

Public Property Let IShape_Colour(ByVal lColour As Long)
    mlColour = lColour
    
    mlRed = lColour And &HFF
    mlGreen = (lColour \ 256) And &HFF
    mlBlue = (lColour \ 65536) And &HFF
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


Public Function IShape_Inside() As Boolean
    Dim oShape As IShape
    
    IShape_Inside = True
    For Each oShape In moShapes
        IShape_Inside = IShape_Inside And oShape.Inside
    Next
    If mbInverse Then
        IShape_Inside = Not IShape_Inside
    End If
End Function

Public Property Let IShape_Inverse(ByVal bInverse As Boolean)
    mbInverse = bInverse
End Property

Public Sub IShape_NextPixel()
    Dim oShape As IShape
    For Each oShape In moShapes
        oShape.NextPixel
    Next
End Sub

Public Sub IShape_NextScanline()
    Dim oShape As IShape
    For Each oShape In moShapes
        oShape.NextScanline
    Next
End Sub

Public Sub IShape_NextSubPixel()
    Dim oShape As IShape
    For Each oShape In moShapes
        oShape.NextSubPixel
    Next
End Sub

Public Sub IShape_NextSubScanline()
    Dim oShape As IShape
    For Each oShape In moShapes
        oShape.NextSubScanline
    Next
End Sub

Public Sub IShape_StartScan()
    Dim oShape As IShape
    For Each oShape In moShapes
        oShape.StartScan
    Next
End Sub
