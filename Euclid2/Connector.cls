VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Public Position As Vector
Public Shapes As Collection

Private moCanvasSurface As Canvas

Private mssShapeState As ShapeStates

Private Sub Class_Initialize()
    Set Position = New Vector
    Set Shapes = New Collection
End Sub

Private Sub IShape_Move(oVector As Vector, ByVal Shift As Integer)
'
End Sub

Public Property Let IShape_State(ByVal ssShapeState As ShapeStates)
    mssShapeState = ssShapeState
    IShape_Render
End Property

Public Property Get IShape_State() As ShapeStates
    IShape_State = mssShapeState
End Property

Public Property Set IShape_CanvasSurface(oCanvasSurface As Canvas)
    Set moCanvasSurface = oCanvasSurface
End Property

Public Property Get IShape_CanvasSurface() As Canvas
    Set IShape_CanvasSurface = moCanvasSurface
End Property

Public Sub IShape_Render()
    Dim lColour As Long
    
    Select Case mssShapeState
        Case ShapeStates.ssNormal
            lColour = euColourNormal
        Case ShapeStates.ssSelected, ShapeStates.ssDragged
            lColour = euColourSelected
        Case ShapeStates.ssRemove
            lColour = moCanvasSurface.Surface.BackColor
    End Select
    
    moCanvasSurface.Surface.Circle (Position.X, Position.Y), TouchRadius, lColour
End Sub

Public Sub IShape_SelectMe(oVector As Vector)
    If mssShapeState <> ssDragged Then
        If oVector.Distance(Position) <= TouchRadius Then
            If mssShapeState <> ssSelected Then
                IShape_State = ssSelected
            End If
        Else
            If mssShapeState = ssSelected Then
                IShape_State = ssNormal
            End If
        End If
    End If
End Sub

Public Sub IShape_DragMe(oVector As Vector)
    If mssShapeState = ssSelected Then
        IShape_State = ssDragged
    End If
End Sub

Public Sub IShape_ReleaseMe()
    IShape_State = ssNormal
End Sub

Public Function IShape_Copy() As IShape
    Dim oShape As IShape
    Dim oConnector As Connector
    
    Set oConnector = New Connector
    
    With oConnector
        Set .IShape_CanvasSurface = moCanvasSurface
        
        .IShape_State = mssShapeState
        
        Set .Position = Position.Clone

        For Each oShape In Shapes
            .Shapes.Add oShape
        Next
    End With
    
    Set IShape_Copy = oConnector
End Function

Public Sub IShape_Translate(oVector As Vector)
    Position.Copy Position.Add(oVector)
End Sub
