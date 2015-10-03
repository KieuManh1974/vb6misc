VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public moResponses As New Collection
Private mdTotalResponseTime As Double
Public SetSize As Long

Public Property Get Average() As Double
    If moResponses.Count > 0 Then
        Average = mdTotalResponseTime / moResponses.Count
    End If
End Property

Public Sub AddResponse(loItem As IGenerator, dResponseTime As Double)
    Dim oStatItem As New clsStatItem
    Dim lIndex As Long
    
    For lIndex = 1 To moResponses.Count
        If moResponses.Item(lIndex).Item.Equals(loItem) Then
            mdTotalResponseTime = mdTotalResponseTime - moResponses.Item(lIndex).ResponseTime
            mdTotalResponseTime = mdTotalResponseTime + dResponseTime
            moResponses.Item(lIndex).ResponseTime = dResponseTime
            moResponses.Item(lIndex).CumulativeResponseTime = moResponses.Item(lIndex).CumulativeResponseTime + dResponseTime
            moResponses.Item(lIndex).Count = moResponses.Item(lIndex).Count + 1
            Exit Sub
        End If
    Next
    
    Set oStatItem.Item = loItem
    oStatItem.ResponseTime = dResponseTime
    oStatItem.CumulativeResponseTime = dResponseTime
    oStatItem.Count = 1
    
    mdTotalResponseTime = mdTotalResponseTime + dResponseTime
    
    moResponses.Add oStatItem
End Sub

Public Function ChooseItem() As IGenerator
    Dim dRandom As Double
    Dim dCumulativeTime As Double
    Dim bOK As Boolean
    Dim vGeneratedItem As Variant
    Dim lIndex As Long
    Dim vItem As Variant
    Dim oGenerator As IGenerator
    
    If moResponses.Count > 0 Then
        dRandom = Rnd * (mdTotalResponseTime * SetSize / moResponses.Count)
        
        For lIndex = 1 To moResponses.Count
            If dCumulativeTime > dRandom Then
                Set ChooseItem = moResponses(lIndex - 1).Item
                Exit Function
            End If
            
            dCumulativeTime = dCumulativeTime + moResponses(lIndex).ResponseTime
        Next
        If dCumulativeTime > dRandom Then
            Set ChooseItem = moResponses(moResponses.Count).Item
            Exit Function
        End If
    End If
    While Not bOK
        Set oGenerator = New clsNumberGenerator
        oGenerator.GenerateItem
        bOK = True
        For lIndex = 1 To moResponses.Count
            If moResponses(lIndex).Item.Equals(oGenerator) Then
                bOK = False
                Exit For
            End If
        Next
    Wend
    Set ChooseItem = oGenerator
End Function
