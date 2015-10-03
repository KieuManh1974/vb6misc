VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThingrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvecList() As New clsVector

Public Property Let Size(ByVal lSize As Long)
    ReDim mvecList(lSize)
End Property

Public Property Set Member(ByVal lIndex As Long, ByVal oVector As clsVector)
    Set mvecList(lIndex) = oVector
End Property

Public Property Get Member(ByVal lIndex As Long) As clsVector
    Set Member = mvecList(lIndex)
End Property

