VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RelationshipList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Function FindRelationship(oFirst As Position, oSecond As Position) As Relationship
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        If oRelationship.FromPos Is oFirst Then
            If oRelationship.ToPos Is oSecond Then
                Set FindRelationship = oRelationship
                Exit Function
            End If
        End If
        If oRelationship.FromPos Is oSecond Then
            If oRelationship.ToPos Is oFirst Then
                Set FindRelationship = oRelationship
                Exit Function
            End If
        End If
    Next
End Function

Public Function RemoveRelationship(oRelationship As Relationship)
    Dim oRelation As Relationship
    Dim iIndex As Long
    
    For iIndex = 1 To List.Count
        If List(iIndex) Is oRelationship Then
            List.Remove iIndex
            Exit Function
        End If
    Next
End Function

Public Function RemoveRelationshipWithReference(oPosition As Position)
    Dim oRelation As Relationship
    Dim bFinished As Boolean
    Dim iIndex As Long
    
    Do
        For iIndex = 1 To List.Count
            If List(iIndex).FromPos Is oPosition Or List(iIndex).ToPos Is oPosition Then
                List.Remove iIndex
                Exit For
            End If
        Next
        If iIndex = List.Count + 1 Then
            bFinished = True
        End If
    Loop While Not bFinished
End Function

Public Function ClearAll()
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        oRelationship.ClearRelationship
    Next
End Function

Public Function RenderAll()
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        oRelationship.RenderRelationship
    Next
End Function
