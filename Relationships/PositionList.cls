VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Function FindPosition(X As Single, Y As Single) As Position
    Dim oPosition As Position
    
    For Each oPosition In List
        If X >= oPosition.LeftTextPosition And X <= oPosition.RightTextPosition Then
            If Y >= oPosition.TopTextPosition And Y <= oPosition.BottomTextPosition Then
                Set FindPosition = oPosition
                Exit For
            End If
        End If
    Next
End Function

Public Function FindPositionFromReference(sRef As String) As Position
    Dim oPosition As Position
    
    For Each oPosition In List
        If oPosition.Reference = sRef Then
            Set FindPositionFromReference = oPosition
            Exit For
        End If
    Next
End Function

Public Function RemovePosition(oPosition As Position)
    Dim iIndex As Long
    
    For iIndex = 1 To List.Count
        If List(iIndex) Is oPosition Then
            List.Remove iIndex
            Exit Function
        End If
    Next
End Function

Public Function RenderAll()
    Dim oPosition As Position
    
    For Each oPosition In List
        oPosition.RenderName
    Next
End Function
