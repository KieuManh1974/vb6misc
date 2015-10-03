VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Public GraticuleH As New clsRational
Public GraticuleV As New clsRational
Public Origin As New clsPoint

Private Sub Class_Initialize()
    GraticuleH.SetRatio 1, 1
    GraticuleV.SetRatio 1, 1
End Sub

Private Function IShape_GetCartesian(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair
    Dim oOrigin As clsCoordinatePair
    
    Dim oCartesian As New clsCoordinatePair
    
    'Set Origin = Origin.Coord.CoordinatePair
    
    oCartesian.SetCoords oPosition.Pair(X), oPosition.Pair(Y)
    
    Set IShape_GetCartesian = oCartesian
    'oOrigin.Pair (X)
End Function

Private Function IShape_GetPolar(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair

End Function

Private Sub IShape_Render(ByVal lColour As Long)
    Dim lIndex As Long
    
    For lIndex = -5 To 5
        
    Next
End Sub
