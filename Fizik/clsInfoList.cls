VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInfoList() As New clsInfo

Public Property Let Size(ByVal lSize As Long)
    ReDim mInfoList(lSize - 1)
End Property

Public Property Set Member(ByVal lIndex As Long, ByVal oThing As clsInfo)
    Set mInfoList(lIndex) = oThing
End Property

Public Property Get Member(ByVal lIndex As Long) As clsInfo
    Set Member = mInfoList(lIndex)
End Property

