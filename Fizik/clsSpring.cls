VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRender
Implements IShape

Private moInfoList As New clsInfoList
Private moDisplayRef As frmDisplay
Private mlShapeType As Long

Private Sub Class_Initialize()
    moInfoList.Size = 1
    mlShapeType = 2
End Sub

Private Property Set IRender_DisplayRef(ByVal oDisplayRef As frmDisplay)
    Set moDisplayRef = oDisplayRef
End Property

Private Sub IRender_Render()
    '
End Sub

Private Property Get IShape_InfoList() As clsInfoList
    Set InfoList = moInfoList
End Property

Private Sub IShape_UpdateMotion(ByVal fInterval As Double)
'
End Sub

Private Property Get IShape_ShapeType() As Long
    IShape_ShapeType = mlShapeType
End Property

