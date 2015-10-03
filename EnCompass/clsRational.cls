VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRational"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mlTerm(1) As Long

Public Property Get Term(ByVal lIndex As Long) As Long
    Term = mlTerm(lIndex)
End Property

Public Property Let Term(ByVal lIndex As Long, ByVal lValue As Long)
    mlTerm(lIndex) = lValue
End Property

Public Property Get DoubleValue() As Double
    DoubleValue = CDbl(mlTerm(0)) / CDbl(mlTerm(1))
End Property

Public Sub SetRatio(ByVal lNumerator As Long, ByVal lDenominator As Long)
    mlTerm(0) = lNumerator
    mlTerm(1) = lDenominator
End Sub
