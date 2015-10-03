VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IShape

Public Origin As New clsCoordinatePair
Public Scaling As Long
Public Canvas As frmCanvas
Public Offset As New clsCoordinatePair

Private Sub Class_Initialize()
    Scaling = 24
    Origin.SetCoords 0, 0
End Sub

Public Function GetPhysicalPosition(ByVal oCoordinatePair As clsCoordinatePair) As clsCoordinatePair
    Set GetPhysicalPosition = New clsCoordinatePair
    GetPhysicalPosition.Pair(X) = Offset.Pair(X) + oCoordinatePair.Pair(X) * Scaling - Origin.Pair(X)
    GetPhysicalPosition.Pair(Y) = Offset.Pair(Y) - oCoordinatePair.Pair(Y) * Scaling - Origin.Pair(Y)
End Function

Public Function GetLogicalPosition(ByVal oCoordinatePair As clsCoordinatePair) As clsCoordinatePair
    Set GetLogicalPosition = New clsCoordinatePair
    GetLogicalPosition.Pair(X) = (oCoordinatePair.Pair(X) + Origin.Pair(X) - Offset.Pair(X)) / Scaling
    GetLogicalPosition.Pair(Y) = (oCoordinatePair.Pair(Y) + Origin.Pair(Y) - Offset.Pair(Y)) / -Scaling
End Function

Public Sub DrawLine(ByVal oOrigin As clsCoordinatePair, ByVal oTarget As clsCoordinatePair, ByVal lColour As Long)
    Dim oRealOrigin As clsCoordinatePair
    Dim oRealTarget As clsCoordinatePair
    
    Set oRealOrigin = GetPhysicalPosition(oOrigin)
    Set oRealTarget = GetPhysicalPosition(oTarget)
    
    Canvas.Line (oRealOrigin.Pair(X), oRealOrigin.Pair(Y))-(oRealTarget.Pair(X), oRealTarget.Pair(Y)), lColour
End Sub

Public Sub DrawCircleFull(ByVal oOrigin As clsCoordinatePair, ByVal fRadius As Double, ByVal lColour As Long)
    Dim oRealOrigin As clsCoordinatePair
    Dim fRealRadius As Double
    
    Set oRealOrigin = GetPhysicalPosition(oOrigin)
    fRealRadius = fRadius * Scaling
    
    Canvas.Circle (oRealOrigin.Pair(X), oRealOrigin.Pair(Y)), fRealRadius, lColour
End Sub

Private Function IShape_GetCartesian(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair
'
End Function

Private Function IShape_GetPolar(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair
'
End Function

Private Sub IShape_Render(ByVal lColour As Long)
'
End Sub
