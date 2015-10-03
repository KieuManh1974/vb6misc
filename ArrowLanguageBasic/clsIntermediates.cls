VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntermediates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moIntermediates() As clsIntermediate

Public Function AddIntermediate(oMember As clsIntermediate)
    ReDim Preserve moIntermediates(Count)
    Set moIntermediates(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Intermediates(ByVal lIndex As Long) As clsIntermediate
    Set Intermediates = moIntermediates(lIndex)
End Property

Public Property Set Intermediates(ByVal lIndex As Long, oMember As clsIntermediate)
    Set moIntermediates(lIndex) = oMember
End Property

