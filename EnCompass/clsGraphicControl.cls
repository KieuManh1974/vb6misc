VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCanvas As frmCanvas
Private moShapeList As New clsShapeList

Public ViewPort As New clsViewPort
Private moCurrentSelection As IShape


Public Property Set Canvas(oCanvas As frmCanvas)
    Set moCanvas = oCanvas
    Set ViewPort.Canvas = oCanvas
End Property

Public Sub MouseAction(ByVal lAction As Long, ByVal oPosition As clsCoordinatePair, ByVal bShift As Boolean, ByVal bControl As Boolean, ByVal bAlt As Boolean)
    Dim oLogicalPos As clsCoordinatePair
    Dim oPhysicalPos As clsCoordinatePair
    
    Dim oSelectionPos As clsCoordinatePair
    
    Select Case lAction
        Case MOUSE_MOVE
            Set oLogicalPos = ViewPort.GetLogicalPosition(oPosition)
            'Set oPhysicalPos = ViewPort.GetPhysicalPosition(oLogicalPos)
            
            Set oSelectionPos = moCurrentSelection.GetCartesian(oLogicalPos)
            
            moCanvas.Caption = oSelectionPos.Pair(X) & "," & oSelectionPos.Pair(Y)
        Case MOUSE_DOWN
        Case MOUSE_UP
    End Select
End Sub

Public Sub Action(ByVal lAction As Long, ByVal oPosition As clsCoordinatePair)
    Select Case lAction
        Case ACTION_MOVE_OFFSET
            Set ViewPort.Offset = oPosition
    End Select
End Sub

Public Sub Test()
    Dim oOrigin As New clsCoordinatePair
    Dim oTarget As New clsCoordinatePair
    
    oOrigin.SetCoords 0, 0
    oTarget.SetCoords 1, 1
    
    ViewPort.DrawLine oOrigin, oTarget, SHOW_NORMAL
    
    oOrigin.SetCoords 1, 1
    oTarget.SetCoords 2, 0
    
    ViewPort.DrawLine oOrigin, oTarget, SHOW_NORMAL
    
    ViewPort.DrawCircleFull oOrigin, 2, SHOW_HIGHLIGHT
    
    Set moCurrentSelection = CreateGrid()
End Sub

Private Function CreateGrid() As clsGrid
    Dim oGrid As New clsGrid
    Dim oPoint As New clsPoint
    
    oPoint.Coord.SetIntegerCoords 0, 0
    Set oGrid.Origin = oPoint
    oGrid.Origin.PointType = POINT_VIEWPORT
    moShapeList.AddShape oGrid
    
    Set CreateGrid = oGrid
End Function

