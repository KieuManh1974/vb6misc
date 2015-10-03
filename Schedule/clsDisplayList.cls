VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisplayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Function FindIdentifierIndex(ByVal lIdentifier As Long) As Long
    Dim oHierarchy As clsHierarchy
    Dim lIndex As Long
    
    For lIndex = 1 To List.Count
        Set oHierarchy = List.Item(lIndex)
        If oHierarchy.Identifier = lIdentifier Then
            FindIdentifierIndex = lIndex - 1
        End If
    Next
End Function
