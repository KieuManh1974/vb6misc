VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVectorSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vVectors() As clsVector3d

Public Sub CreateSet(ParamArray vSet() As Variant)
    Dim lIndex As Long
    
    ReDim vVectors(UBound(vSet))
    
    For lIndex = 0 To UBound(vSet)
        vVectors = vSet(lIndex)
    Next
End Sub

Public Property Get Vector(lIndex As Long) As clsVector3d
    Set Vector = vVectors(lIndex)
End Property

