VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumberGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IGenerator

Private mvItem As Variant

Private Function IGenerator_Answer() As Variant
    
End Function

Private Function IGenerator_CheckAnswer(mvAnswer As Variant) As Boolean
    Dim lAnswer As Long
    
    lAnswer = (mvItem(0) + mvItem(1)) Mod 10
    
    IGenerator_CheckAnswer = (mvAnswer = lAnswer)
End Function

Private Function IGenerator_Equals(loItem As IGenerator) As Boolean
    Dim vItem As Variant
    
    vItem = loItem.GetItem
    If vItem(0) = mvItem(0) Then
        If vItem(1) = mvItem(1) Then
            IGenerator_Equals = True
        End If
    End If
End Function

Private Sub IGenerator_GenerateItem()
    mvItem = Array(Int(Rnd * 9), Int(Rnd * 9))
End Sub

Private Function IGenerator_GetItem() As Variant
    IGenerator_GetItem = mvItem
End Function
