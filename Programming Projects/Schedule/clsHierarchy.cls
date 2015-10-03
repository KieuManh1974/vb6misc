VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHierarchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Task As New clsTask

Public Level As Long
Public Expanded As Boolean

Public Children As New Collection
Public Parent As clsHierarchy

Public Identifier As Long

Private Const Statuses = "ACH"

Public Property Let Descriptor(sDescriptor As String)
    Task.Status = InStr(Statuses, Left$(sDescriptor, 1)) - 1
    Task.Description = Mid$(sDescriptor, 5)
    
    Level = Val(Mid$(sDescriptor, 2, 2))
    Expanded = Mid$(sDescriptor, 4, 1) = "+"
End Property

Public Property Get Descriptor() As String
    Descriptor = Mid$(Statuses, Task.Status + 1, 1) & Format$(Level, "00") & IIf(Expanded, "+", "-") & Task.Description
End Property

Public Property Get Display() As String
    'Display = IIf(Task.Completed, "-- ", IIf(Task.Hold, "** ", "")) & String$(Level * 4, " ") & IIf(Not Expanded And Children.Count > 0, "+", "") & Task.Description
End Property

Private Sub Class_Initialize()
    Level = -1
    Expanded = True
End Sub

Public Function FindIdentifierIndex(ByVal lIdentifier As Long) As Long
    Dim oHierarchy As clsHierarchy
    Dim lIndex As Long
    
    For lIndex = 1 To Children.Count
        Set oHierarchy = Children.Item(lIndex)
        If oHierarchy.Identifier = lIdentifier Then
            FindIdentifierIndex = lIndex
            Exit Function
        End If
    Next
End Function

