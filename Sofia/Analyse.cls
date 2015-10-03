VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Analyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum StatementTypes
    stShortVar
    stLongVar
    stLoop
    stExpression
    stFunction
    stBlock
    stConditional
End Enum

Public Sub CompileParseTree(oBlock As SaffronTree, oContext As clsNode)
    Dim oList As SaffronTree
    Dim lIndex As Long
    Dim lFunctionSize As Long
    Dim sFunctionName As String
    Dim bIsFunction As Boolean
    
    Set oList = oBlock(1)
    
    lIndex = 1
    Do
        If oList(lIndex).Index = 2 Then
            oList.SubTree.Remove lIndex
            lIndex = lIndex - 1
        End If
        lIndex = lIndex + 1
    Loop Until lIndex > oList.SubTree.Count
    
    lIndex = 1
    
    ' Handle functions
    Select Case LCase$(oList(lIndex).Text)
        Case "function"
            sFunctionName = oList(lIndex + 1).Text
            oContext.TextKey = sFunctionName
            oContext.Value = Array(stFunction, sFunctionName)
            lIndex = lIndex + 2
        Case "loop"
            oContext.Value = Array(stLoop, "")
            lIndex = lIndex + 1
        Case "if"
            oContext.Value = Array(stConditional, "")
            lIndex = lIndex + 1
        Case Else
            oContext.Value = Array(stBlock, "")
    End Select
    
    ' Handle vars and expressions
    While lIndex <= oList.SubTree.Count
        Select Case oList(lIndex).Index
            Case 1 ' identifier
                Select Case LCase$(oList(lIndex).Text)
                    Case "short"
                        oContext.AddNew().Value = Array(stShortVar, oList(lIndex + 1).Text)
                        lIndex = lIndex + 2
                    Case "long"
                        oContext.AddNew().Value = Array(stLongVar, oList(lIndex + 1).Text)
                        lIndex = lIndex + 2
                    Case Else 'expression
                        oContext.AddNew().Value = Array(stExpression, oList(lIndex).Text)
                        lIndex = lIndex + 1
                End Select
            Case 3 ' sub block {}
                CompileParseTree oList(lIndex)(1), oContext.AddNew
                lIndex = lIndex + 1
        End Select
    Wend
End Sub
