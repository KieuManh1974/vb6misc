VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Expression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ReturnType As Container
Private ExecutionStack As New Collection
Public Parent As Object

Public Function Compile(oTree As ParseTree)

End Function

Public Function CompileOld(oTree As ParseTree, oScope As Container) As Container
    Dim lIndex As Long
    Dim oPart As ParseTree
    Dim oNewContext As Block
    Dim sIdentifier As String
    Dim stScope As ScopeTypes
    
    Dim oBlock As Block
    Dim oFunction As Block
    
    Set oTree = oTree(1)
    
    lIndex = 1
    Set oPart = oTree.SubTree(lIndex)
    sIdentifier = oPart.Text
    stScope = stAny
    
    While Not oScope Is Nothing
        ' Numeric
        
        ' Variables
        
        ' Blocks
        Set oBlock = oScope.FindLongest(sIdentifier, oScope.Blocks, stScope)
        If Not oBlock Is Nothing Then
            stScope = stPublic
            sIdentifier = Right$(sIdentifier, Len(sIdentifier) - Len(oBlock.Name))
            Set oScope = oBlock
            If sIdentifier = "" Then
                lIndex = lIndex + 1
                If lIndex <= oTree.SubTree.Count Then
                    Set oPart = oTree.SubTree(lIndex)
                    sIdentifier = oPart.Text
                Else
                    ExecutionStack.Add oBlock
                    Set CompileExpression = oBlock
                    Exit Function
                End If
            End If
        End If
        
        ' Functions
        Set oFunction = oScope.FindLongest(sIdentifier, oScope.Functions, stScope)
        If Not oFunction Is Nothing Then
            stScope = stPublic
            sIdentifier = Right$(sIdentifier, Len(sIdentifier) - Len(oFunction.Name))
            Set oScope = oFunction.ReturnType
            
            If sIdentifier = "" Then
                lIndex = lIndex + 1
                If lIndex <= oTree.SubTree.Count Then
                    Set oPart = oTree.SubTree(lIndex)
                    sIdentifier = oPart.Text
                Else
                End If
            End If
        End If

    Wend
End Function

Public Function Execute()

End Function
