VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Solid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oPoints As New Points
Private vJoins As Variant
Private pi2 As Double
Public Paper As Form

Public Sub Construct(iSides As Integer)
    Dim oYaxis As New Vector
    Dim oXAxis As New Vector
    
    oYaxis.Create 0, Cos(2), Sin(2)
    oXAxis.Create Cos(3), Sin(3), 0
    
    vJoins = Array()
    pi2 = Atn(1) * 8
    ConstructBase2 iSides, oXAxis, oYaxis
End Sub

Public Sub Render()
    Dim vJoin As Variant
    
    For Each vJoin In vJoins
        Paper.Line (oPoints.Points(vJoin(0)).Flat(0), oPoints.Points(vJoin(0)).Flat(1))-(oPoints.Points(vJoin(1)).Flat(0), oPoints.Points(vJoin(1)).Flat(1)), vbWhite
    Next
End Sub


Private Sub ConstructBase2(iSides As Integer, oVector1 As Vector, ovector2 As Vector)
    Dim oXAxis As New Vector
    Dim oYaxis As New Vector
    Dim iSide As Integer
    Dim oPoint As Vector
    
    Dim ov1 As Vector
    Dim ov1a As Vector
        
    Dim ov2 As Vector
    Dim ov2a As Vector
    Dim on1 As Vector
    
    Dim dRadius As Double
    
    dRadius = 100 * Sin(pi2 / iSides / 2)
    
    Dim vFacePoints As Variant
    vFacePoints = Array()
    
    Set oXAxis = oVector1
    Set oYaxis = oVector1.Cross(ovector2).Cross(oVector1)
    
    For iSide = 0 To iSides - 1
        Set oPoint = oXAxis.Scalar(Cos(pi2 * iSide / iSides) * dRadius).Add(oYaxis.Scalar(Sin(pi2 * iSide / iSides) * dRadius))
        ReDim Preserve vFacePoints(iSide) As Variant
        vFacePoints(iSide) = oPoints.AddPoint(oPoint)
    Next
    
    For iSide = 0 To iSides - 1
        AddJoin vFacePoints(iSide), vFacePoints((iSide + 1) Mod iSides)
        Set ov1 = oPoints.VectorOfPoints(vFacePoints((iSide - 1 + iSides) Mod iSides), vFacePoints(iSide))
        Set ov2 = oPoints.VectorOfPoints(vFacePoints((iSide + 1) Mod iSides), vFacePoints(iSide))
        
        Set on1 = ov2.Cross(ov1).Scalar(100).Add(oPoints.Points(vFacePoints(iSide)))
        AddJoin vFacePoints(iSide), oPoints.AddPoint(on1)
    Next
    
End Sub

Private Function AddJoin(ByVal lPoint1 As Long, ByVal lPoint2 As Long) As Long
    ReDim Preserve vJoins(UBound(vJoins) + 1) As Variant
    AddJoin = UBound(vJoins)
    vJoins(AddJoin) = Array(lPoint1, lPoint2)
End Function

