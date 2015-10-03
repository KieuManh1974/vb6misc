VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Compose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlBlockCounter As Long

Private Enum StatementTypes
    stShortVar
    stLongVar
    stLoop
    stExpression
    stFunction
    stBlock
    stConditional
End Enum

Public Function CompileBlocks(oAnalysed As clsNode) As Block
    Dim oItem As clsNode
    Dim lIndex As Long
    Dim oBlock As New Block
    Dim oSubBlock As Block
    
    For lIndex = 0 To oAnalysed.Count - 1
        Set oItem = oAnalysed.ItemPhysical(lIndex)
        
        Select Case oItem.Value(0)
            Case stShortVar
                oBlock.moVariables.AddNew(, oItem.Value(1)).Value = 1
            Case stLongVar
                oBlock.moVariables.AddNew(, oItem.Value(1)).Value = 2
            Case stLoop
                Set oSubBlock = CompileBlocks(oItem)
                Set oSubBlock.moParentBlock = oBlock
                oSubBlock.msName = NewBlockLabel
                oSubBlock.btBlockType = btLoop
                Set oBlock.moIntermediateCode.AddNew().Value = oSubBlock
            Case stBlock
                Set oSubBlock = CompileBlocks(oItem)
                Set oSubBlock.moParentBlock = oBlock
                oSubBlock.msName = NewBlockLabel
                oSubBlock.btBlockType = btBlock
                Set oBlock.moIntermediateCode.AddNew().Value = oSubBlock
            Case stConditional
                Set oSubBlock = CompileBlocks(oItem)
                Set oSubBlock.moParentBlock = oBlock
                oSubBlock.msName = NewBlockLabel
                oSubBlock.btBlockType = btConditional
                Set oBlock.moIntermediateCode.AddNew().Value = oSubBlock
            Case stFunction
                Set oSubBlock = CompileBlocks(oItem)
                Set oSubBlock.moParentBlock = oBlock
                oSubBlock.msName = oItem.Value(1)
                oSubBlock.btBlockType = btFunction
                Set oBlock.moSubBlocks.AddNew(, oSubBlock.msName).Value = oSubBlock
            Case stExpression
                oBlock.moIntermediateCode.AddNew().Value = oItem.Value(1)
        End Select
    Next
    
    Set CompileBlocks = oBlock
End Function

Private Function NewBlockLabel() As String
    NewBlockLabel = "block" & mlBlockCounter
    mlBlockCounter = mlBlockCounter + 1
End Function


