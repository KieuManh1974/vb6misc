VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Identifier As String
Public Name As String
Public SourceID As Long
Public Aliases As Variant
Public msAlias As String

Public Property Get UniqueName() As String
    UniqueName = Name
End Property

Public Property Get DisplayName() As String
    DisplayName = Name
End Property

Public Property Let Alias(ByVal sAlias As String)
    msAlias = sAlias
End Property

Public Property Get Alias() As String
    If msAlias <> "" Then
        Alias = msAlias
    Else
        Alias = Name
    End If
End Property
