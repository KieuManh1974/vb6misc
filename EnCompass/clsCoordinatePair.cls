VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoordinatePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private moPair(1) As Double

Public Property Get Pair(ByVal lIndex As Long) As Double
    Pair = moPair(lIndex)
End Property

Public Property Let Pair(ByVal lIndex As Long, ByVal fValue As Double)
    moPair(lIndex) = fValue
End Property

Public Function SetCoords(ByVal fValue0 As Double, ByVal fValue1 As Double)
    moPair(0) = fValue0
    moPair(1) = fValue1
End Function
