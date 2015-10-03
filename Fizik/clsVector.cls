VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mfVector(1) As Double

Public Property Get Component(ByVal lIndex As Long) As Double
    Component = mfVector(lIndex)
End Property

Public Property Let Component(ByVal lIndex As Long, ByVal fValue As Double)
    mfVector(lIndex) = fValue
End Property

Public Sub SetComponents(ByVal fValue0 As Double, ByVal fvalue1 As Double)
    mfVector(0) = fValue0
    mfVector(1) = fvalue1
End Sub

Public Function Add(ByVal oVector As clsVector) As clsVector
    Dim lIndex As Long
    Dim oNew As New clsVector
    
    For lIndex = 0 To 1
        oNew.Component(lIndex) = mfVector(lIndex) + oVector.Component(lIndex)
    Next
    
    Set Add = oNew
End Function

Public Function Subt(ByVal oVector As clsVector) As clsVector
    Dim lIndex As Long
    Dim oNew As New clsVector
    
    For lIndex = 0 To 1
        oNew.Component(lIndex) = mfVector(lIndex) - oVector.Component(lIndex)
    Next
    
    Set Subt = oNew
End Function
