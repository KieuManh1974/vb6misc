VERSION 5.00
Begin VB.Form clsFineDiagram 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   115
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "clsFineDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public moShapeStack As New Collection


Private Sub Form_Load()
    ScaleTop = ScaleHeight
    ScaleHeight = -ScaleHeight
End Sub

Public Sub Render()
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
    
            StartCounter
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
                            lRed = lRed + oShape.Red
                            lGreen = lGreen + oShape.Green
                            lBlue = lBlue + oShape.Blue
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
            SetPixelV hdc, lX, -ScaleHeight - lY, RGB(lRed \ 256, lGreen \ 256, lBlue \ 256)
            For Each oShape In moShapeStack
                oShape.NextPixel
            Next
        Next
        For Each oShape In moShapeStack
            oShape.NextScanline
        Next
    Next
            Debug.Print GetCounter
    
End Sub

Public Sub AddFilledCircle(oCentre As clsVector, nRadius As Single, lColour As Long)
    Dim oCircle As New clsCircle
    
    oCircle.SetUp oCentre, nRadius
    oCircle.IShape_Colour = lColour

    moShapeStack.Add oCircle
End Sub

Public Sub AddFilledSector(oCentre As clsVector, nRadius As Single, nStartAngle As Single, nEndAngle As Single, lColour As Long)
    Dim oCircle1 As New clsCircle
    Dim oTriangle1 As New clsTriangle
    Dim oV As New clsVector
    Dim oComposite As New clsCompositeAnd
    
    oCircle1.SetUp oCentre, nRadius
    oTriangle1.SetUp oCentre, oV.Create(Sqr(2) * nRadius * Cos(nStartAngle), Sqr(2) * nRadius * Sin(nStartAngle)).Add(oCentre), oV.Create(Sqr(2) * nRadius * Cos(nEndAngle), Sqr(2) * nRadius * Sin(nEndAngle)).Add(oCentre)
    oTriangle1.IShape_Inverse = True
    oComposite.AddShape oCircle1
    oComposite.AddShape oTriangle1
    oComposite.IShape_Colour = lColour
    moShapeStack.Add oComposite
End Sub

Public Sub AddSector(oCentre As clsVector, nRadius As Single, nThickness As Single, nStartAngle As Single, nEndAngle As Single, lColour As Long)
    Dim oCircle1 As New clsCircle
    Dim oCircle2 As New clsCircle
    Dim oTriangle1 As New clsTriangle
    Dim oV As New clsVector
    Dim oComposite As New clsCompositeAnd
    
    oCircle1.SetUp oCentre, nRadius + nThickness
    oCircle2.SetUp oCentre, nRadius - nThickness
    oCircle2.IShape_Inverse = True
    oTriangle1.SetUp oCentre, oV.Create(Sqr(2) * (nRadius + nThickness) * Cos(nStartAngle), Sqr(2) * (nRadius + nThickness) * Sin(nStartAngle)).Add(oCentre), oV.Create(Sqr(2) * (nRadius + nThickness) * Cos(nEndAngle), Sqr(2) * (nRadius + nThickness) * Sin(nEndAngle)).Add(oCentre)
    oComposite.AddShape oCircle1
    oComposite.AddShape oCircle2
    oComposite.AddShape oTriangle1
    
    oComposite.IShape_Colour = lColour
    moShapeStack.Add oComposite
End Sub

Public Sub AddCircle(oCentre As clsVector, nRadius As Single, nThickness As Single, lColour As Long)
    Dim oCircle1 As New clsCircle
    Dim oCircle2 As New clsCircle
    Dim oComposite As New clsCompositeAnd
    
    oCircle1.SetUp oCentre, nRadius + nThickness
    oCircle2.SetUp oCentre, nRadius - nThickness
    oCircle2.IShape_Inverse = True
    
    oComposite.AddShape oCircle1
    oComposite.AddShape oCircle2
    
    oComposite.IShape_Colour = lColour
    
    moShapeStack.Add oComposite
End Sub

Public Sub AddFilledTriangle(oPoint1 As clsVector, oPoint2 As clsVector, oPoint3 As clsVector, lColour As Long)
    Dim oTriangle As New clsTriangle
    
    oTriangle.SetUp oPoint1, oPoint2, oPoint3
    oTriangle.IShape_Colour = lColour
    moShapeStack.Add oTriangle
End Sub

Public Sub AddLine(oPoint1 As clsVector, oPoint2 As clsVector, nThickness As Single, lColour As Long)
    Dim oBase As clsVector
    Dim oPerp As clsVector
    Dim oTriangle1 As New clsTriangle
    Dim oTriangle2 As New clsTriangle
    
    Set oBase = oPoint2.Subs(oPoint1).Perpendicular.Normal.Scalar(nThickness)
    oTriangle1.SetUp oPoint1.Add(oBase), oPoint1.Subs(oBase), oPoint2.Add(oBase)
    oTriangle1.IShape_Colour = lColour
    oTriangle2.SetUp oPoint1.Subs(oBase), oPoint2.Subs(oBase), oPoint2.Add(oBase)
    oTriangle2.IShape_Colour = lColour
    moShapeStack.Add oTriangle1
    moShapeStack.Add oTriangle2
End Sub

Public Sub AddBox(oPoint1 As clsVector, oPoint2 As clsVector, lColour As Long)
    Dim oTriangle1 As New clsTriangle
    Dim oTriangle2 As New clsTriangle
    Dim v As New clsVector
    
    Set oPoint2 = oPoint1.Add(oPoint2)
    oTriangle1.SetUp oPoint1, v.Create(oPoint1.x, oPoint2.y), oPoint2
    oTriangle1.IShape_Colour = lColour
    oTriangle2.SetUp oPoint1, v.Create(oPoint2.x, oPoint1.y), oPoint2
    oTriangle2.IShape_Colour = lColour
    
    moShapeStack.Add oTriangle1
    moShapeStack.Add oTriangle2
End Sub

Public Sub AddRoundBox(oPoint1 As clsVector, oPoint2 As clsVector, nRadius As Single, lColour As Long)
    Dim v As New clsVector
    
    AddBox v.Create(oPoint1.x + nRadius, oPoint1.y), v.Create(oPoint2.x - nRadius, oPoint2.y), lColour
    AddBox v.Create(oPoint1.x, oPoint1.y + nRadius), v.Create(oPoint2.x, oPoint2.y - nRadius), lColour
    AddFilledCircle v.Create(oPoint1.x + nRadius, oPoint1.y + nRadius), nRadius, lColour
    AddFilledCircle v.Create(oPoint1.x + nRadius, oPoint2.y - nRadius), nRadius, lColour
    AddFilledCircle v.Create(oPoint2.x - nRadius, oPoint2.y - nRadius), nRadius, lColour
    AddFilledCircle v.Create(oPoint2.x - nRadius, oPoint1.y + nRadius), nRadius, lColour
End Sub

Public Sub AddGridBoxes(oTopLeft As clsVector, oBoxDimensions As clsVector, oSeparation As clsVector, oSize As clsVector, lColour As Long)
    Dim nx As Single
    Dim ny As Single
    Dim v As New clsVector
    Dim oBoxCorner As New clsVector
    
    For ny = 1 To oSize.y
        For nx = 1 To oSize.x
            oBoxCorner.SetVector (nx - 1) * oSeparation.x + oTopLeft.x, (ny - 1) * oSeparation.y + oTopLeft.y
            AddBox oBoxCorner, oBoxCorner.Add(oBoxDimensions), lColour
        Next
    Next
End Sub

Public Sub AddGridCircles(oMidPoint As clsVector, oSeparation As clsVector, nRadius As Single, oSize As clsVector, lColour As Long)
    Dim nx As Single
    Dim ny As Single
    Dim v As New clsVector
    Dim oBoxCorner As New clsVector
    

    oBoxCorner.x = oMidPoint.x - ((oSize.x - 1) / 2) * oSeparation.x
    oBoxCorner.y = oMidPoint.y - ((oSize.x - 1) / 2) * oSeparation.y
    
    For ny = 1 To oSize.y
        For nx = 1 To oSize.x
            AddFilledCircle v.Create(oBoxCorner.x + oSeparation.x * (nx - 1), oBoxCorner.y + oSeparation.y * (ny - 1)), nRadius, lColour
        Next
    Next
End Sub

Public Sub AddArrowTriangle(oMidPoint As clsVector, oDimensions As clsVector, sAngle As Single, lColour As Long)
    Dim oTip As clsVector
    Dim v As New clsVector
    Dim nRealAngle As Single
    Dim oShaft As clsVector
    Dim oBackLeft As clsVector
    Dim oBackRight As clsVector
    
    nRealAngle = Atn(1) * 8 * sAngle
    Set oShaft = v.Create(oDimensions.x * Cos(sRealAngle) / 2, oDimensions.x * Sin(sRealAngle) / 2)
    Set oTip = oMidPoint.Add(oShaft)
    Set oBackLeft = oMidPoint.Subs(oShaft).Add(oShaft.Perpendicular.Normal.Scalar(oDimensions.y / 2))
    Set oBackRight = oMidPoint.Subs(oShaft).Subs(oShaft.Perpendicular.Normal.Scalar(oDimensions.y / 2))
    
    AddFilledTriangle oTip, oBackLeft, oBackRight, lColour
End Sub


Public Sub AddChevron(oMidPoint As clsVector, oDimensions As clsVector, nThickness As Single, sAngle As Single, lColour As Long)
    Dim oXAxis As clsVector
    Dim oYAxis As clsVector
    Dim v As New clsVector
    Dim nRealAngle As Single
    Dim oApex As clsVector
    Dim oBase As clsVector
    Dim oLeftBase As clsVector
    Dim oRightBase As clsVector
    Dim oLeftBaseTop As clsVector
    Dim oRightBaseTop As clsVector
    
    nRealAngle = Atn(1) * 8 * sAngle
    Set oXAxis = v.Create(Cos(nRealAngle), Sin(nRealAngle))
    Set oYAxis = v.Create(-Sin(nRealAngle), Cos(nRealAngle))
    
    Set oApex = oXAxis.Scalar(nThickness)
    Set oBase = oXAxis.Scalar(nThickness - oDimensions.x)
    Set oLeftBase = oBase.Add(oYAxis.Scalar(oDimensions.y / 2))
    Set oRightBase = oBase.Add(oYAxis.Scalar(-oDimensions.y / 2))
    Set oLeftBaseTop = oLeftBase.Add(oApex)
    Set oRightBaseTop = oRightBase.Add(oApex)
    
    AddFilledTriangle oMidPoint, oLeftBase.Add(oMidPoint), oLeftBaseTop.Add(oMidPoint), lColour
    AddFilledTriangle oLeftBaseTop.Add(oMidPoint), oMidPoint, oApex.Add(oMidPoint), lColour
    AddFilledTriangle oMidPoint, oRightBase.Add(oMidPoint), oRightBaseTop.Add(oMidPoint), lColour
    AddFilledTriangle oRightBaseTop.Add(oMidPoint), oMidPoint, oApex.Add(oMidPoint), lColour
End Sub
