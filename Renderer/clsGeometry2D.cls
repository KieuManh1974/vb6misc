VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeometry2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function VectorCrossesPlane(vsPhoton As clsVectorSet, vsTriangle As clsVectorSet) As clsVector3d
    Dim v1 As clsVector3d
    Dim v2 As clsVector3d
    Dim vN As clsVector3d
    Dim vS As clsVector3d
    Dim dE As Double
    Dim vQ As clsVector3d
    Dim vR As clsVector3d
    Dim vP As clsVector3d
    Dim vT As clsVector3d
    Dim vCrossPoint As clsVector3d
    
    vP = vsTriangle.Vector(0)
    
    v1 = vsTriangle.Vector(1).Subs(vP)
    v2 = vsTriangle.Vector(2).Subs(vP)
    vN = v1.Cross(v2)
    
    vQ = vsPhoton.Vector(0)
    vR = vsPhoton.Vector(1)
    vT = vQ.Subs(vP)
    
    dE = vT.Dot(vN) / vR.Dot(vN)
    
    vCrossPoint = vQ.Add(vR.Scalar(dE))
End Function

'Public Function PointInPolygon(vPoint As clsVector3d, pPolygon As clsPolygon) As Boolean
'
'End Function
