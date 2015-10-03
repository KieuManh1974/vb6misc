VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Top As clsHierarchy

Public Sub RefreshLevels()
    RefreshLevel Top
End Sub

Private Sub RefreshLevel(oParent As clsHierarchy, Optional ByVal lLevel As Long = -1)
    Dim oHierarchy As clsHierarchy
    
    oParent.Level = lLevel
    For Each oHierarchy In oParent.Children
        RefreshLevel oHierarchy, lLevel + 1
    Next
End Sub
