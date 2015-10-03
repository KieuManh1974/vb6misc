VERSION 5.00
Begin VB.Form clsDiagram 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "clsDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public moShapeStack As New Collection

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Sub Form_Load()
    ScaleTop = ScaleHeight
    ScaleHeight = -ScaleHeight
End Sub

Public Sub Render()
    Dim oShape As Object
    
    FillStyle = vbSolid

    For Each oShape In moShapeStack
        If TypeOf oShape Is clsCircle Then
            FillColor = oShape.IShape_Colour
            Circle (oShape.Centre.x, oShape.Centre.y), oShape.Radius, oShape.IShape_Colour
        ElseIf TypeOf oShape Is clsTriangle Then
            Triangle oShape.Point(1), oShape.Point(2), oShape.Point(3), oShape.IShape_Colour
        End If
    Next
End Sub

Private Sub Triangle(oPoint1 As clsVector, oPoint2 As clsVector, oPoint3 As clsVector, ByVal lColour As Long)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(2) As POINTAPI
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(hdc, hPen)
    hOldBrush = SelectObject(hdc, hBrush)
        
    paPoints(0).x = oPoint1.x
    paPoints(0).y = oPoint1.y
    paPoints(1).x = oPoint2.x
    paPoints(1).y = oPoint2.x
    paPoints(2).x = oPoint3.x
    paPoints(2).y = oPoint3.x
    
    Polygon hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(hdc, hOldPen)
    Call SelectObject(hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub


Public Sub AddFilledCircle(oCentre As clsVector, lRadius As Long, lColour As Long)
    Dim oCircle As New clsCircle
    
    oCircle.SetUp oCentre, lRadius
    oCircle.IShape_Colour = lColour

    moShapeStack.Add oCircle
End Sub

Public Sub AddCircle(oCentre As clsVector, nRadius As Single, nThickness As Single, lColour As Long)
    Dim oCircle1 As New clsCircle
    Dim oCircle2 As New clsCircle
    Dim oComposite As New clsCompositeAnd
    
    oCircle1.SetUp oCentre, nRadius + nThickness
    oCircle1.IShape_Colour = lColour
    oCircle2.SetUp oCentre, nRadius - nThickness
    oCircle2.IShape_Colour = vbBlack
   
    moShapeStack.Add oCircle1
    moShapeStack.Add oCircle2
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

