VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte

Private moShapeStack As New Collection

Private Sub Form_Activate()
    DrawShapes
End Sub

Private Sub Form_Load()
    ScaleTop = ScaleHeight
    ScaleHeight = -ScaleHeight

    InitialiseShapes
End Sub

Private Sub InitialiseShapes()
    Dim lRandom As Long
    Dim lPoint As Long
    Dim oPoints(2) As clsVector
    Dim oVector As clsVector
    Dim oShape As clsTriangle
    Dim oCircle As clsCircle
    
    Randomize
    
    Set oVector = New clsVector
    oVector.SetVector 70, 70
    Set oCircle = New clsCircle
    oCircle.SetUp oVector, 50
    oCircle.IShape_Colour = vbGreen
    moShapeStack.Add oCircle

    Set oVector = New clsVector
    oVector.SetVector 70, 70
    Set oCircle = New clsCircle
    oCircle.SetUp oVector, 40
    'oCircle.IShape_Inverse = True
    oCircle.IShape_Colour = vbBlack
    moShapeStack.Add oCircle

    For lRandom = 1 To 1
        For lPoint = 0 To 2
            Set oVector = New clsVector
            oVector.SetVector Int(Rnd * 200), Int(Rnd * 150)
            Set oPoints(lPoint) = oVector
        Next
        Set oShape = New clsTriangle
        oShape.SetUp oPoints(0), oPoints(1), oPoints(2)
        oShape.IShape_Colour = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        moShapeStack.Add oShape
    Next
End Sub

Private Sub DrawShapes()
    Dim oShape As IShape
    Dim lShapeIndex As Long
    Dim lY As Long
    Dim lX As Long
    Dim lSubX As Long
    Dim lSubY As Long
    Dim lRed As Double
    Dim lGreen As Double
    Dim lBlue As Double
    
    For Each oShape In moShapeStack
        oShape.StartScan
    Next
    
    For lY = 0 To -ScaleHeight
        For lX = 0 To ScaleWidth
            lRed = 0
            lGreen = 0
            lBlue = 0
            For lSubY = 0 To 15
                For lSubX = 0 To 15
                    For lShapeIndex = moShapeStack.Count To 1 Step -1
                        Set oShape = moShapeStack(lShapeIndex)
        
                        If oShape.Inside Then
                            lRed = lRed + (oShape.Colour And &HFF)
                            lGreen = lGreen + ((oShape.Colour \ 256) And &HFF)
                            lBlue = lBlue + ((oShape.Colour \ 65536) And &HFF)
                            Exit For
                        End If
                    Next
                    For Each oShape In moShapeStack
                        oShape.NextSubPixel
                    Next
                Next
                For Each oShape In moShapeStack
                    oShape.NextSubScanline
                Next
            Next
            SetPixelV hDC, lX, -ScaleHeight - lY, RGB(lRed \ 256, lGreen \ 256, lBlue \ 256)
            For Each oShape In moShapeStack
                oShape.NextPixel
            Next
        Next
        For Each oShape In moShapeStack
            oShape.NextScanline
        Next
    Next
End Sub

