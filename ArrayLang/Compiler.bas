Attribute VB_Name = "Compiler"
Option Explicit

Private moVariables As clsNode
Private moTypes As clsNode

Public Sub InitialiseCompiler()
    Set moVariables = New clsNode
    Set moTypes = New clsNode
End Sub

Public Sub CompileProgram()
    Dim oTree As SaffronTree
    
    SaffronStream.Text = Definition.msProgram
    Set oTree = New SaffronTree
    
    If Definition.moParser.Parse(oTree) Then
        Debug.Print Program(oTree)
    Else
        MsgBox "Compilation Error"
        End
    End If
End Sub

Private Function Program(oTree As SaffronTree) As String
    Dim oSubTree As SaffronTree
    
    For Each oSubTree In oTree.SubTree
        Select Case oSubTree.Index
            Case 1 ' declaration
                Declaration oSubTree(1)
            Case 2 ' type def
                TypeDef oSubTree(1)
            Case 3 ' copy
                Copy oSubTree(1)
        End Select
    Next
End Function

Private Function Range(oTree As SaffronTree) As Variant
    Range = Array(Val(oTree(1).Text), Val(oTree(2).Text))
End Function

Private Function RangeSet(oTree As SaffronTree) As Variant
    Dim oSubTree As SaffronTree
    Dim vRangeSet As Variant
    Dim lIndex As Long
    
    vRangeSet = Array()
    ReDim vRangeSet(oTree(1).SubTree.Count - 1)
    
    For Each oSubTree In oTree(1).SubTree
        vRangeSet(lIndex) = Range(oSubTree)
        lIndex = lIndex + 1
    Next
    RangeSet = vRangeSet
End Function

Private Sub TypeDef(oTree As SaffronTree)
    Dim sName As String
    Dim vRangeSet As Variant
    Dim oType As clsNode
    
    sName = oTree(1).Text
    vRangeSet = RangeSet(oTree(2))
    
    Set oType = moTypes.AddNew(, sName)
    oType.Value = vRangeSet
End Sub

Private Sub Declaration(oTree As SaffronTree)
    Dim oType As clsNode
    Dim oVariable As clsNode
    Dim vRangeSet As Variant
    Dim sName As String
    Dim oVariableInfo As clsNode
    Dim oVariableInfoType As clsNode
    
    Select Case oTree.Index
        Case 1 ' Range Name
            vRangeSet = RangeSet(oTree(1)(1))
            Set oType = moTypes.AddNew()
            oType.Value = vRangeSet
            
            sName = oTree(1)(2).Text
            Set oVariable = moVariables.AddNew(, sName)
            Set oVariableInfo = New clsNode
            Set oVariableInfoType = oVariableInfo.AddNew(, "type")
            Set oVariableInfoType.Value = oType
            Set oVariable.Value = oVariableInfo
        Case 2 ' Name Name
            sName = oTree(1)(1).Text
            Set oType = moTypes.Keys.Word(sName).Value
            
            sName = oTree(1)(2).Text
            Set oVariable = moVariables.AddNew(, sName)
            Set oVariableInfo = New clsNode
            Set oVariableInfoType = oVariableInfo.AddNew(, "type")
            Set oVariableInfoType.Value = oType
            Set oVariable.Value = oVariableInfo
    End Select
End Sub

Private Sub Copy(oTree As SaffronTree)
    Dim sVariable As String
    Dim oVariable As clsNode
    Dim vValue As Variant
    Dim oVariableInfo As clsNode
    Dim oVariableInfoValue As clsNode
    
    sVariable = oTree(1).Text
    Set oVariable = moVariables.Keys.Word(sVariable).Value
    
    Set oVariableInfo = oVariable.Value.Keys.Word("type").Value
    Set oVariableInfoValue = oVariable.Value.AddNew(, "value")
    oVariableInfoValue.Value = SetValue(oTree(2))
End Sub

Private Function SetValue(oTree As SaffronTree) As Variant
    Dim oSubTree As SaffronTree
    Dim vSetValue As Variant
    Dim lIndex As Long
    
    vSetValue = Array()
    ReDim vSetValue(oTree.SubTree.Count)
    
    For Each oSubTree In oTree(1).SubTree
        Select Case oSubTree.Index
            Case 1 ' number
                vSetValue(lIndex) = Val(oSubTree(1).Text)
            Case 2 ' variable
                vSetValue(lIndex) = oSubTree(1).Text
            Case 3 ' set
                vSetValue(lIndex) = SetValue(oSubTree(1))
        End Select
        lIndex = lIndex + 1
    Next
    SetValue = vSetValue
End Function
