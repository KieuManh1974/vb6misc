VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRatCoordinatePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private moPair(1) As New clsRational

Public Property Get Pair(ByVal lIndex As Long) As clsRational
    Set Pair = moPair(lIndex)
End Property

Public Property Set Pair(ByVal lIndex As Long, ByVal oRational As clsRational)
    Set moPair(lIndex) = oRational
End Property

Public Property Get CoordinatePair() As clsCoordinatePair
    Set CoordinatePair = New clsCoordinatePair
    CoordinatePair.Pair(0) = moPair(0).DoubleValue
    CoordinatePair.Pair(1) = moPair(1).DoubleValue
End Property

Public Function SetIntegerCoords(ByVal lX As Long, ByVal lY As Long)
    moPair(0).SetRatio lX, 1
    moPair(1).SetRatio lY, 1
End Function
