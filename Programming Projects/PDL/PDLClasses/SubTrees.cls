VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubTrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oItems As New Collection

Public Sub Add(oTree As ParseTree)
    Dim lIndex As Long
    Dim sKey As String
    
    sKey = oTree.Name
    If sKey = "" Then
        oItems.Add oTree
    Else
        On Error Resume Next
        Do
            If oItems(sKey & IIf(lIndex = 0, "", lIndex)) Is Nothing Then
                Exit Do
            End If
                
            lIndex = lIndex + 1
        Loop
        oItems.Add oTree, sKey & IIf(lIndex = 0, "", lIndex)
    End If
End Sub

Public Property Get Item(ByVal vIndex As Variant) As ParseTree
Attribute Item.VB_UserMemId = 0
    Set Item = oItems(vIndex)
End Property

Public Property Get Count() As Long
    Count = oItems.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = oItems.[_NewEnum]
End Property


