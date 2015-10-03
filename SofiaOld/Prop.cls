VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Scope As ScopeTypes
Public Name As String
Public ContainerRefs As New Collection

Public Containers As New Collection
Public Properties As New Collection
Public Code As New Collection
Public Parent As Object
Public ExpressionParser As String


Public Sub Compile(oTree As ParseTree)
    Dim oStatement As ParseTree
    Dim oContainer As Container
    Dim oProperty As Prop
    Dim oFlow As Flow
    Dim oExpression As Expression
    
    Select Case oTree(1).Index
        Case 0 ' No scope
            Scope = stPrivate
        Case 1 ' Scope
            Select Case oTree(1)(1).Index
                Case 1 ' public
                    Scope = stPublic
                Case 2 ' private
                    Scope = stPrivate
            End Select
    End Select
    
    Name = oTree(2).Text
    
    For Each oStatement In oTree(3)(1).SubTree
        Select Case oStatement.Index
            Case 1 ' Block
                Select Case oStatement(1).Index
                    Case 2 ' Container
                        Set oContainer = New Container
                        oContainer.Compile oStatement(1)(1)
                        
                    Case 3 ' Property
                        Set oProperty = New Prop
                        oProperty.Compile oStatement(1)(1)
                End Select
        End Select
    Next
    
    For Each oStatement In oTree(1).SubTree
        Select Case oStatement.Index
            Case 1 ' Block
                Select Case oStatement(1).Index
                    Case 1 ' Flow
                        Set oFlow = New Flow
                End Select
    
            Case 2 ' Expression
                Set oExpression = New Expression
        End Select
    Next
        
End Sub


Public Function CompileOld(oTree As ParseTree)
    Dim oCode As ParseTree
    'Dim oBlock As Block
    Dim oExpression As Expression
    Dim oBlockInfo As ParseTree
    
    For Each oCode In oTree(1).SubTree
        Select Case oCode.Index
            Case 1 ' Block def
                'Set oBlock = New Block
                
                Select Case LCase$(oCode(1)(1).Text)
                    Case "private"
                        oBlock.Scope = stPrivate
                    Case "public"
                        oBlock.Scope = stPublic
                    Case Else
                        oBlock.Scope = stPrivate
                End Select
                
                Set oBlockInfo = oCode(1)(2)
                Select Case oBlockInfo.Index
                    Case 1 ' Block
                        oBlock.Name = oBlockInfo(1).Text
                    Case Else ' Function
                        Set oBlock.ReturnType = FindBlock(oBlockInfo(oBlockInfo.SubTree.Count - 1).Text)
                        oBlock.Name = oBlockInfo(oBlockInfo.SubTree.Count).Text
                End Select
                
                Set oBlock.Parent = Me
                
                oBlock.CompileBlock oCode(1)(3)
                If Not oBlock.ReturnType Is Nothing Then
                    Functions.Add oBlock
                Else
                    Blocks.Add oBlock
                End If
                
'            Case 2 ' Expression
'                Set oExpression = New Expression
'                oExpression.CompileExpression oCode(1)(1), Me
'                Expressions.Add oExpression
        End Select
    Next
    
    ExpressionParser = CreateParserExpression()
End Function

Public Function CreateParserExpression() As String
    Dim oBlock As Block
    Dim sSubParser As String
    
    For Each oBlock In Blocks
        If oBlock.ExpressionParser <> "" Then
            sSubParser = "(AND '" & oBlock.Name & "', ? " & oBlock.ExpressionParser & ")"
        Else
            sSubParser = "'" & oBlock.Name & "'"
        End If
        CreateParserExpression = CreateParserExpression & ", " & sSubParser
    Next
    If CreateParserExpression <> "" Then
        CreateParserExpression = Mid$(CreateParserExpression, 3)
        If Blocks.Count > 1 Then
            CreateParserExpression = "(OR " & CreateParserExpression & ")"
        End If
    End If
End Function

Public Function FindBlock(sBlockName As String) As Container
    Dim oBlock As Block
    
    For Each oBlock In Blocks
        If oBlock.Name = sBlockName Then
            Set FindBlock = oBlock
            Exit Function
        End If
    Next
    
    If FindBlock Is Nothing Then
        Set FindBlock = Parent.FindBlock(sBlockName)
    End If
End Function

Public Function FindLongest(sName As String, oCollection As Collection, stScope As ScopeTypes) As Object
    Dim lLength As Long
    Dim sSubName As String
    Dim oObject As Object
    
    For lLength = Len(sName) To 1 Step -1
        sSubName = Left$(sName, lLength)
        For Each oObject In oCollection
            If stScope = stAny Or stScope = oObject.Scope Then
                If oObject.Name = sSubName Then
                    Set FindLongest = oObject
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Public Function Execute()

End Function


