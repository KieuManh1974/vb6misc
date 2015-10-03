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
Public Sub Render(ByVal lColour As Long)

End Sub

Public Function GetCartesian(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair

End Function

Public Function GetPolar(ByVal oPosition As clsCoordinatePair) As clsCoordinatePair

End Function
