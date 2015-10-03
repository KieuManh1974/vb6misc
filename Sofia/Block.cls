VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BlockTypes
    btNone
    btFunction
    btLoop
    btBlock
    btConditional
End Enum

Public btBlockType As BlockTypes
Public msName As String
Public moParentBlock As Block
Public moVariables As New clsNode
Public moSubBlocks As New clsNode
Public moIntermediateCode As New clsNode


