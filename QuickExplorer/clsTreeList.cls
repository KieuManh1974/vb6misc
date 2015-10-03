VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moTreeList As New Collection

Public Function NewNode() As clsTreeNode
    Set NewNode = New clsTreeNode
    moTreeList.Add NewNode
End Function

Public Function RemoveNode(oNode As clsTreeNode)
    Dim oTreeNode As clsTreeNode
    Dim lIndex As Long
    
    For Each oTreeNode In moTreeList
        lIndex = lIndex + 1
        If oTreeNode Is oNode Then
            moTreeList.Remove lIndex
            Exit Function
        End If
    Next
End Function

Public Function FindNodeAtPosition(ByVal lXpos As Long, ByVal lYPos As Long) As Variant
    Dim oTreeNode As clsTreeNode
    Dim lIndex As Long
    Dim stSelection As SelectionType
    
    FindNodeAtPosition = Array(Nothing, stNone)
    For Each oTreeNode In moTreeList
        stSelection = oTreeNode.InRectangle(lXpos, lYPos)
        If stSelection <> stNone Then
            FindNodeAtPosition = Array(oTreeNode, stSelection)
            Exit Function
        End If
    Next
End Function
