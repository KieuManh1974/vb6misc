VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get State() As ShapeStates
'
End Property

Public Property Let State(ByVal ssShapeState As ShapeStates)
'
End Property

Public Sub Render()
'
End Sub

Public Sub SelectMe(oVector As Vector)
'
End Sub

Public Sub DragMe(oVector As Vector)
'
End Sub

Public Sub Translate(oVector As Vector)
'
End Sub

Public Sub Move(oVector As Vector, ByVal Shift As Integer)
'
End Sub

Public Sub ReleaseMe()
'
End Sub

Public Function Copy() As IShape
'
End Function

Public Property Get CanvasSurface() As Canvas
'
End Property

Public Property Set CanvasSurface(oCanvasSurface As Canvas)
'
End Property

Private Property Let ShapeState(ByVal ssShapeState As ShapeStates)
'
End Property

Private Property Get ShapeState() As ShapeStates
'
End Property
