VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public ExpressionType As ExpressionTypes
Public Identifiers As New Collection
Private msExpression As String
Private moTree As ParseTree

Public Enum ExpressionTypes
    etSelect
    etWhere
End Enum

Public Property Get UniqueIdentifier() As String
    If moTree Is Nothing Then
        UniqueIdentifier = msExpression
    Else
        UniqueIdentifier = moTree.Text
    End If
End Property

Public Property Get Aliased() As String
    Dim oPart As ParseTree
    Dim lIdentifierIndex As Long
    Dim oIdentifier As clsIdentifier
    Dim oTable As clsTableInfo
    
    lIdentifierIndex = 1
    
    For Each oPart In moTree.SubTree
        Aliased = Aliased & oPart(1).Text
        If oPart(2).Index = 1 Then
            Set oIdentifier = Identifiers.Item(lIdentifierIndex)
            If oIdentifier.Table <> "ANY" Then
                Set oTable = goTables.FindTable(oIdentifier.Table, oIdentifier.Database)
                Aliased = Aliased & oTable.Alias & "." & oIdentifier.Field
            Else
                Set oTable = goFields.SelectedTable
                Aliased = Aliased & oTable.Alias & "." & oIdentifier.Field
            End If
            lIdentifierIndex = lIdentifierIndex + 1
        End If
    Next
End Property

Public Property Let Expression(sString As String)
    Dim oTree As New ParseTree
    Dim oPart As ParseTree
    Dim oSubPart As ParseTree
    Dim oIdentifier As clsIdentifier
    
    Stream.Text = sString
    If goParseExpression.Parse(oTree) Then
        Set moTree = oTree
        For Each oPart In moTree.SubTree
            If oPart(2).Index = 1 Then
                Set oIdentifier = New clsIdentifier
                Set oIdentifier.IdentifierTree = oPart(2)(1)
                Identifiers.Add oIdentifier
            End If
        Next
    Else
        msExpression = sString
    End If
End Property


