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
    Dim lPositionIndex As Long
    
    For lPositionIndex = List.Count To 1 Step -1
        Set oPosition = List(lPositionIndex)
        If X >= (oPosition.PosX - oPosition.Radius) And X <= (oPosition.PosX + oPosition.Radius) Then
            If Y >= (oPosition.PosY - oPosition.Radius) And Y <= (oPosition.PosY + oPosition.Radius) Then
                Set FindPosition = oPosition
                Exit For
            End If
        End If
    Next
End Function

Public Function FindPositions(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Collection
    Dim oPosition As Position
    Dim MidX As Single
    Dim MidY As Single
    Dim WidthX As Single
    Dim WidthY As Single
    Dim nTemp As Single
    
    If X2 < X1 Then
        nTemp = X1
        X1 = X2
        X2 = nTemp
    End If
    
    If Y2 < Y1 Then
        nTemp = Y1
        Y1 = Y2
        Y2 = nTemp
    End If
    
    Set FindPositions = New Collection
    
    For Each oPosition In List
        With oPosition
        
        If X1 <= oPosition.PosX And X2 >= oPosition.PosX Then
            If Y1 <= oPosition.PosY And Y2 >= oPosition.PosY Then
                    FindPositions.Add oPosition
                End If
            End If
        End With
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
