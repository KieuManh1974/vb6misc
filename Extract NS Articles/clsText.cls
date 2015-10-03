VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Found As Boolean
Public Position As Long
Public PositionSearch As Long
Public Text As String
Public TextSearch As String

Public msText As String

Public Function Find(ByVal sSearch, ByVal lPosition As Long) As Variant
    Dim lFoundPos As Long
    
    lFoundPos = InStr(lPosition, msText, sSearch)
    Find = Array(lFoundPos <> 0, lFoundPos)
End Function

Public Function FindNext(ByVal lPosition As Long, ParamArray vSearch() As Variant) As Variant
    Dim lClosest As Long
    Dim vItem As Variant
    Dim vFind As Variant
    Dim bFound As Boolean
    Dim vClosestItem As Variant
    
    vClosestItem = ""
    lClosest = Len(msText) + 1
    For Each vItem In vSearch
        vFind = Find(vItem, lPosition)
        If vFind(0) Then
            If vFind(1) < lClosest Then
                lClosest = vFind(1)
                vClosestItem = vItem
                bFound = True
            End If
        End If
    Next
    FindNext = Array(bFound, lClosest + Len(vClosestItem), Mid$(msText, lPosition, lClosest - lPosition), lClosest, vClosestItem)
End Function

Public Sub FindNext2(ByVal lPosition As Long, ParamArray vSearch() As Variant)
    Dim lClosest As Long
    Dim vItem As Variant
    Dim vFind As Variant
    Dim bFound As Boolean
    Dim vClosestItem As Variant
    
    vClosestItem = ""
    lClosest = Len(msText) + 1
    For Each vItem In vSearch
        vFind = Find(vItem, lPosition)
        If vFind(0) Then
            If vFind(1) < lClosest Then
                lClosest = vFind(1)
                vClosestItem = vItem
                bFound = True
            End If
        End If
    Next
       
    Found = bFound
    Position = lClosest + Len(vClosestItem)
    PositionSearch = lClosest
    Text = Mid$(msText, lPosition, lClosest - lPosition)
    TextSearch = vClosestItem
End Sub
