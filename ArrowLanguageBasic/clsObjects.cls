VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moObjects() As clsObject

Public Function AddObject(oMember As clsObject)
    ReDim Preserve moObjects(Count)
    Set moObjects(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Objects(ByVal lIndex As Long) As clsObject
    Set Objects = moObjects(lIndex)
End Property

Public Property Set Objects(ByVal lIndex As Long, oMember As clsObject)
    Set moObjects(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String) As clsObject
    Dim lIndex As Long
    
    For lIndex = 0 To Count - 1
        If moObjects(lIndex).Identifier = sIdentifier Then
            Set GetByIdentifier = moObjects(lIndex)
            Exit Function
        End If
    Next
End Function

