VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public moStrs As New clsStrs
Public moRecs As New clsRecs
Public moVars As New clsVars

Public moParentContext As clsContext

Public moStatements As clsStatements

Public Function FindVarByName(ByVal sIdentifier As String) As clsVar
    Dim oVar As clsVar
    Dim lIndex As Long
    
    For lIndex = 0 To moVars.Count - 1
        Set oVar = moVars.Vars(lIndex)
        If oVar.Identifier = sIdentifier Then
            Set FindVarByName = oVar
            Exit Function
        End If
    Next
End Function
