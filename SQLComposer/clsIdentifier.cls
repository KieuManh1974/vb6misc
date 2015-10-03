VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msDatabase As String
Private msTable As String
Private msField As String
Private msIdentifier As String

Public Property Get Database() As String
    Database = msDatabase
End Property

Public Property Get Table() As String
    Table = msTable
End Property

Public Property Get Field() As String
    Field = msField
End Property

Public Property Set IdentifierTree(oTree As ParseTree)
    Select Case oTree.Index
        Case 2
            msDatabase = oTree(1).Text
            msTable = oTree(2).Text
            msIdentifier = msDatabase & "." & msTable
        Case 3
            msDatabase = oTree(1).Text
            msTable = oTree(2).Text
            msField = oTree(3).Text
            msIdentifier = msDatabase & "." & msTable & "." & msField
    End Select
End Property

Public Property Let Identifier(ByVal sIdentifier As String)
    Dim oTree As New ParseTree
    
    Stream.Text = sIdentifier
    
    If goParseIdentifier.Parse(oTree) Then
        Set IdentifierTree = oTree
    Else
        msIdentifier = sIdentifier
    End If
End Property

Public Property Get Identifier() As String
    Identifier = msIdentifier
End Property
