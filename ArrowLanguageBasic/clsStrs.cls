VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moStrs() As clsStr

Public Function AddStr(oMember As clsStr)
    ReDim Preserve moStrs(Count)
    Set moStrs(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Strs(ByVal lIndex As Long) As clsStr
    Set Strs = moStrs(lIndex)
End Property

Public Property Set Strs(ByVal lIndex As Long, oMember As clsStr)
    Set moStrs(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String) As clsStr
    Dim lIndex As Long
    
    For lIndex = 0 To Count - 1
        If moStrs(lIndex).Identifier = sIdentifier Then
            Set GetByIdentifier = moStrs(lIndex)
            Exit Function
        End If
    Next
End Function


