VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Scope As ScopeTypes
Public Name As String

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
                        Set oContainer.Parent = Me
                        Containers.Add oContainer
                        
                    Case 3 ' Property
                        Set oProperty = New Prop
                        oProperty.Compile oStatement(1)(1)
                        Set oProperty.Parent = Me
                        Properties.Add oProperty
                End Select
        End Select
    Next
    
    CreateParserExpression
    
    For Each oStatement In oTree(3)(1).SubTree
        Select Case oStatement.Index
            Case 1 ' Block
                Select Case oStatement(1).Index
                    Case 1 ' Flow
                        Set oFlow = New Flow
                        oFlow.Compile oStatement(1)(1)
                        Set oFlow.Parent = Me
                        Code.Add oFlow
                End Select
    
            Case 2 ' Expression
                Set oExpression = New Expression
                oExpression.Compile oStatement(1)(1)
                Set oExpression.Parent = Me
                Code.Add oExpression
        End Select
    Next
        
End Sub

Private Sub CreateParserExpression()
    Dim oParent As Object
    
    Set oParent = Me
    
    While Not oParent Is Nothing
        ExpressionParser = "_" & oParent.Name & ExpressionParser
        Set oParent = oParent.Parent
    Wend
    
    ExpressionParser = Mid$(ExpressionParser, 2) & " := "
    
    
End Sub


'Public Function CreateParserExpression() As String
'    Dim oBlock As Block
'    Dim sSubParser As String
'
'    For Each oBlock In Blocks
'        If oBlock.ExpressionParser <> "" Then
'            sSubParser = "(AND '" & oBlock.Name & "', ? " & oBlock.ExpressionParser & ")"
'        Else
'            sSubParser = "'" & oBlock.Name & "'"
'        End If
'        CreateParserExpression = CreateParserExpression & ", " & sSubParser
'    Next
'    If CreateParserExpression <> "" Then
'        CreateParserExpression = Mid$(CreateParserExpression, 3)
'        If Blocks.Count > 1 Then
'            CreateParserExpression = "(OR " & CreateParserExpression & ")"
'        End If
'    End If
'End Function

'Public Function FindBlock(sBlockName As String) As Container
'    Dim oBlock As Block
'
'    For Each oBlock In Blocks
'        If oBlock.Name = sBlockName Then
'            Set FindBlock = oBlock
'            Exit Function
'        End If
'    Next
'
'    If FindBlock Is Nothing Then
'        Set FindBlock = Parent.FindBlock(sBlockName)
'    End If
'End Function
'
'Public Function FindLongest(sName As String, oCollection As Collection, stScope As ScopeTypes) As Object
'    Dim lLength As Long
'    Dim sSubName As String
'    Dim oObject As Object
'
'    For lLength = Len(sName) To 1 Step -1
'        sSubName = Left$(sName, lLength)
'        For Each oObject In oCollection
'            If stScope = stAny Or stScope = oObject.Scope Then
'                If oObject.Name = sSubName Then
'                    Set FindLongest = oObject
'                    Exit Function
'                End If
'            End If
'        Next
'    Next
'End Function

'Public Function Execute()
'
'End Function

