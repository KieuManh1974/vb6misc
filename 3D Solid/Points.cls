VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Points"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPoints As Variant

Public Property Get Points(ByVal iIndex As Long) As Vector
    Set Points = vPoints(iIndex)
End Property

Public Function VectorOfPoints(ByVal lPoint1 As Long, ByVal lPoint2 As Long) As Vector
    Dim oP1 As Vector
    Dim oP2 As Vector
    Set oP1 = vPoints(lPoint2)
    Set oP2 = vPoints(lPoint1)
    Set VectorOfPoints = oP2.Subs(oP1)
End Function

Public Function AddPoint(oPoint As Vector) As Long
    Dim vPoint As Vector
    Dim iPointIndex As Integer
    
    For iPointIndex = 0 To UBound(vPoints)
        If vPoints(iPointIndex).Subs(oPoint).Size < 2 Then
            AddPoint = iPointIndex
            Exit Function
        End If
    Next

    ReDim Preserve vPoints(UBound(vPoints) + 1) As Variant
    AddPoint = UBound(vPoints)
    Set vPoints(AddPoint) = oPoint
End Function

Private Sub Class_Initialize()
    vPoints = Array()
End Sub
