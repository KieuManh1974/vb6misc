VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mlStrs() As clsStr
Private mlVars() As clsVar
Private mlRecs() As clsRec

Public Function AddStr(oStr As clsStr)
    ReDim Preserve moStrs(Count)
    Set moStrs(Count) = oStr
    Count = Count + 1
End Function

Public Property Get Strs(ByVal lIndex As Long) As clsStr
    Set Strs = moStrs(lIndex)
End Property

Public Property Set Strs(ByVal lIndex As Long, oStr As clsStr)
    Set moStrs(lIndex) = oStr
End Property



Public Function AddVar(oVar As clsVar)
    ReDim Preserve moVars(Count)
    Set moVars(Count) = oVar
    Count = Count + 1
End Function

Public Property Get Vars(ByVal lIndex As Long) As clsVar
    Set Vars = moVars(lIndex)
End Property

Public Property Set Vars(ByVal lIndex As Long, oVar As clsVar)
    Set moVars(lIndex) = oVar
End Property


Public Function AddRec(oRec As clsRec)
    ReDim Preserve moRecs(Count)
    Set moRecs(Count) = oRec
    Count = Count + 1
End Function

Public Property Get Recs(ByVal lIndex As Long) As clsRec
    Set Recs = moRecs(lIndex)
End Property

Public Property Set Recs(ByVal lIndex As Long, oRec As clsRec)
    Set moRecs(lIndex) = oRec
End Property
