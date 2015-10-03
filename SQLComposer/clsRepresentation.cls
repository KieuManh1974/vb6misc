VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DisplayedTables As New Dictionary
Public DisplayedFields As New Dictionary
Public DisplayedSQL As New clsSQL
Public SelectedTables As New Dictionary
Public SelectedFields As New Dictionary

Public Function CreateTable(sName As String, lSourceID As Long, sSourceShortName As String, sAliases As Variant, sAlias As String) As clsTableInfo
    Dim oTable As New clsTableInfo
    
    With oTable
        .Name = sName
        .SourceID = lSourceID
        .SourceShortName = sSourceShortName
        .Aliases = Split(sAliases, " ")
    End With
End Function

Public Function CreateField(sName As String, oTable As clsTableInfo) As clsFieldInfo
    Dim oField As New clsFieldInfo
    
    With oField
        .Name = sName
        Set .TableInfo = oTable
    End With
End Function

Public Sub AddTable(oTable As clsTableInfo)
    If Not DisplayedTables.Exists(oTable.UniqueName) Then
        DisplayedTables.Add oTable.UniqueName, oTable
    End If
End Sub

Public Sub AddField(oField As clsFieldInfo)
    If Not DisplayedFields.Exists(oField.UniqueName) Then
        DisplayedFields.Add oField.UniqueName, oField
    End If
End Sub

Public Sub RemoveTable(oTable As clsTableInfo)
    If DisplayedTables.Exists(oTable.UniqueName) Then
        DisplayedTables.Remove oTable.UniqueName
    End If
End Sub

Public Sub RemoveField(oField As clsFieldInfo)
    If DisplayedFields.Exists(oField.UniqueName) Then
        DisplayedFields.Remove oField.UniqueName
    End If
End Sub

