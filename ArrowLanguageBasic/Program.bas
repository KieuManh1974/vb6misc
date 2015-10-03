Attribute VB_Name = "Program"
Option Explicit

'Private ProgramAnalyse As New Analyse
'Private ProgramCompose As New Compose
'Private ProgramGenerate As New GenerateIntermediateCode

Private Types As New Collection
Private Vars As New Collection

Public Function LoadProgram() As clsContext
    Dim oFSO As New FileSystemObject
    Dim oTree As New SaffronTree
    Dim oContext As New clsContext
    
    SaffronStream.Text = oFSO.OpenTextFile(App.Path & "\testprogram.txt", ForReading).ReadAll
    Set oTree = New SaffronTree
    If Not oParser.Parse(oTree) Then
        MsgBox "Syntax error"
        End
    End If
    
    DoStatements oTree, oContext
    
    Set LoadProgram = oContext
End Function

Private Sub DoStatements(oTree As SaffronTree, oContext As clsContext)
    Dim oStatement As SaffronTree
    
    For Each oStatement In oTree.SubTree
        DoStatement oStatement, oContext
    Next
End Sub

Private Sub DoStatement(oTree As SaffronTree, oContext As clsContext)
    Select Case oTree.Index
        Case 1 ' Class Declaration
            DoClassDeclaration oTree(1), oContext
        Case 2 ' Object Instance
            DoObject oTree(1), oContext
        Case 3 ' Expression
            DoExpression oTree(1), oContext
    End Select
End Sub

Private Sub DoClassDeclaration(oTree As SaffronTree, oContext As clsContext)
    Dim oClass As New clsClass
    
    oClass.Identifier = oTree(1).Text
    oClass.Range.Starting = oTree(2)(1)(1).Text
    oClass.Range.Ending = oTree(2)(1)(2).Text
    
    oContext.moClasses.AddClass oClass
End Sub

Private Sub DoObject(oTree As SaffronTree, oContext As clsContext)
    Dim oObject As New clsObject
    Dim oClass As clsClass
    
    oObject.Identifier = oTree(1).Text
    
    Select Case oTree(2).Index
        Case 1 ' class reference
            Set oObject.UnitClass = oContext.moClasses.GetByIdentifier(oTree(2)(1).Text)
        Case 2 ' straight unit
            Set oObject.UnitClass = New clsClass
            oObject.UnitClass.Range.Starting = oTree(2)(1)(1)(1).Text
            oObject.UnitClass.Range.Ending = oTree(2)(1)(1)(2).Text
    End Select
    
    oContext.moObjects.AddObject oObject
End Sub

Private Function TempVarLabel(sIdentify As String) As String
    Dim lIndex As Long
    
    For lIndex = 1 To 5
        TempVarLabel = TempVarLabel & Chr$(Rnd() * 26 + 65)
    Next
    TempVarLabel = sIdentify & "_" & TempVarLabel
End Function

Private Function DoExpression(oTree As SaffronTree, oContext As clsContext) As clsObject
    Dim bTerm As Boolean
    Dim oSubTree As SaffronTree
    Dim oPreviousObject As clsObject
    Dim oObject As clsObject
    Dim opOperator As ILOPERATORS
    Dim oIntermediate As clsIntermediate
    Dim oTempObject As clsObject
    Dim oTempClass As clsClass
    Dim oConstantObject As clsObject
    
    bTerm = True
    
    For Each oSubTree In oTree.SubTree
        Select Case bTerm
            Case True ' Term
                If oPreviousObject Is Nothing Then
                    Select Case oSubTree.Index
                        Case 1 ' identifier
                            Set oPreviousObject = oContext.moObjects.GetByIdentifier(oSubTree.Text)
                            Set DoExpression = oPreviousObject
                        Case 2 ' constant
                            Set oObject = New clsObject
                            oObject.Identifier = TempVarLabel("const")
                            oObject.IsConstant = True
                            Set oTempClass = New clsClass
                            oTempClass.Range.Starting = oSubTree(1).Text
                            oTempClass.Range.Ending = oSubTree(1).Text
                            Set oObject.UnitClass = oTempClass
                            oContext.moObjects.AddObject oObject
                            Set oPreviousObject = oObject
                            Set DoExpression = oPreviousObject
                        Case 3 ' brackets
                            Set oPreviousObject = DoExpression(oSubTree(1)(1), oContext)
                            Set DoExpression = oPreviousObject
                    End Select
                Else
                    Select Case oSubTree.Index
                        Case 1 ' identifier
                            Set oObject = oContext.moObjects.GetByIdentifier(oSubTree.Text)
                        Case 2 ' constant
                            Set oObject = New clsObject
                            oObject.Identifier = TempVarLabel("const")
                            oObject.IsConstant = True
                            Set oTempClass = New clsClass
                            oTempClass.Range.Starting = oSubTree(1).Text
                            oTempClass.Range.Ending = oSubTree(1).Text
                            Set oObject.UnitClass = oTempClass
                            oContext.moObjects.AddObject oObject
                        Case 3 ' brackets
                            Set oObject = DoExpression(oSubTree(1)(1), oContext)
                    End Select
                    Set oIntermediate = New clsIntermediate
                    oIntermediate.Operator = opOperator
                    Set oIntermediate.Operand1 = oPreviousObject
                    Set oIntermediate.Operand2 = oObject
                    Set oPreviousObject = oPreviousObject
                    Set DoExpression = oPreviousObject
                    Select Case opOperator
                        Case opAdd, opSub, opMultiply, opDivide, opModulus
                            Set oTempClass = New clsClass
                            
                            Select Case opOperator
                                Case opAdd
                                    oTempClass.Range.Starting = oIntermediate.Operand1.UnitClass.Range.Starting + oIntermediate.Operand2.UnitClass.Range.Starting
                                    oTempClass.Range.Ending = oIntermediate.Operand1.UnitClass.Range.Ending + oIntermediate.Operand2.UnitClass.Range.Ending
                                Case opSub
                                    oTempClass.Range.Starting = oIntermediate.Operand1.UnitClass.Range.Starting - oIntermediate.Operand2.UnitClass.Range.Ending
                                    oTempClass.Range.Ending = oIntermediate.Operand1.UnitClass.Range.Ending - oIntermediate.Operand2.UnitClass.Range.Starting
                                Case opMultiply
                                    oTempClass.Range.Starting = oIntermediate.Operand1.UnitClass.Range.Starting * oIntermediate.Operand2.UnitClass.Range.Starting
                                    oTempClass.Range.Ending = oIntermediate.Operand1.UnitClass.Range.Ending * oIntermediate.Operand2.UnitClass.Range.Ending
                                Case opDivide
                                    oTempClass.Range.Starting = Int(oIntermediate.Operand1.UnitClass.Range.Starting / oIntermediate.Operand2.UnitClass.Range.Ending)
                                    oTempClass.Range.Ending = Int(oIntermediate.Operand1.UnitClass.Range.Ending / oIntermediate.Operand2.UnitClass.Range.Starting)
                                Case opModulus
                                    oTempClass.Range.Starting = 0
                                    oTempClass.Range.Ending = oIntermediate.Operand2.UnitClass.Range.Ending - 1
                            End Select

                            Set oTempObject = New clsObject
                            Set oTempObject.UnitClass = oTempClass
                            oTempObject.Identifier = TempVarLabel("temp")
                            oContext.moObjects.AddObject oTempObject
                            Set oIntermediate.Operand3 = oTempObject
                            Set oPreviousObject = oTempObject
                            Set DoExpression = oPreviousObject
                    End Select
                    oContext.moIntermediates.AddIntermediate oIntermediate
                End If
                bTerm = Not bTerm
            Case False ' Operator
                Select Case oSubTree.Index
                    Case 1 ' copy
                        opOperator = opCopy
                    Case 2 ' add
                        opOperator = opAdd
                    Case 3 ' subtract
                        opOperator = opSub
                    Case 4 ' multiply
                        opOperator = opMultiply
                    Case 5 ' divide
                        opOperator = opDivide
                    Case 6 ' modulus
                        opOperator = opModulus
                End Select
                bTerm = Not bTerm
        End Select
    Next

End Function

'Private Sub DoDeclaration(oTree As SaffronTree, oContext As clsContext)
'    Select Case oTree(1).Index
'        Case 1 ' rec
'            DoRec oTree, oContext
'        Case 2 ' str
'            DoStr oTree, oContext
'        Case 3 ' var
'            DoVar oTree, oContext
'    End Select
'End Sub

'Private Sub DoRec(oTree As SaffronTree, oContext As clsContext)
'    Dim oRec As New clsRec
'
'    oRec.Identifier = oTree(2).Text
'    Set oRec.Extent = DoExtent(oTree(3), oContext)
'
'    oContext.moRecs.AddRec oRec
'End Sub
'
'Private Sub DoStr(oTree As SaffronTree, oContext As clsContext)
'    Dim oStr As New clsStr
'
'    oStr.Identifier = oTree(2).Text
'    Set oStr.Extent = DoExtent(oTree(3), oContext)
'
'    oContext.moStrs.AddStr oStr
'End Sub
'
'Private Sub DoVar(oTree As SaffronTree, oContext As clsContext)
'    Dim oVar As New clsVar
'
'    oVar.Identifier = oTree(2).Text
'    Set oVar.VarArray.Extents = DoExtent(oTree(3), oContext)
'
'    oContext.moVars.AddVar oVar
'End Sub
'
'Private Function DoExtent(oTree As SaffronTree, oContext As clsContext) As clsExtents
'    Dim oExtent As clsExtent
'    Dim oExtents As New clsExtents
'    Dim oSubTree As SaffronTree
'    Dim oSet As Collection
'    Dim oMember As clsIMember
'    Dim lIndex As Long
'
'    lIndex = 1
'    For Each oSubTree In oTree.SubTree
'        Select Case oSubTree.Index
'            Case 1 ' index set
'                Set oExtent = New clsExtent
'                Set oSet = DoIndexSet(oSubTree)
'                For Each oMember In oSet
'                    oExtent.AddMember oMember
'                Next
'                oExtents.AddExtent oExtent
'
'            Case 2 ' identifier ' first must be a rec, others must be str
'                Select Case lIndex
'                    Case 1
'                        Set oExtent = New clsExtent
'                        oExtent.AddMember oContext.moRecs.GetByIdentifier(oSubTree.Text)
'                        oExtents.AddExtent oExtent
'                    Case 2
'                        Set oExtent = New clsExtent
'                        oExtent.AddMember oContext.moStrs.GetByIdentifier(oSubTree.Text)
'                        oExtents.AddExtent oExtent
'                End Select
'        End Select
'        lIndex = lIndex + 1
'    Next
'
'    Set DoExtent = oExtents
'End Function
'
'Private Function DoIndexSet(oTree As SaffronTree) As Collection
'    Dim oSubTree As SaffronTree
'    Dim oRangeSet As clsRangeSet
'    Dim oSymbolSet As clsSymbolSet
'    Dim oIdentifierSet As clsIdentifierSet
'    Dim sIdentifier As String
'    Dim oIdentifier As SaffronTree
'    Dim oMember As clsIMember
'
'    Set DoIndexSet = New Collection
'
'    For Each oSubTree In oTree(1)(1).SubTree
'        sIdentifier = oSubTree(1).Text
'
'        Select Case oSubTree(2).Index
'            Case 1 ' symbol set
'                Set oSymbolSet = New clsSymbolSet
'                oSymbolSet.SymbolSet = oSubTree(2).Text
'                Set oMember = oSymbolSet
'                oMember.Identifier = sIdentifier
'            Case 2 ' identifier set
'                Set oIdentifierSet = New clsIdentifierSet
'                For Each oIdentifier In oSubTree(2)(1).SubTree
'                    oIdentifierSet.AddIdentifier oIdentifier.Text
'                Next
'                Set oMember = oIdentifierSet
'                oMember.Identifier = sIdentifier
'            Case 3 ' range set
'                Set oRangeSet = DoRangeSet(oSubTree(2)(1))
'                Set oMember = oRangeSet
'                oMember.Identifier = sIdentifier
'        End Select
'
'        DoIndexSet.Add oMember
'    Next
'End Function
'
'Private Function DoRangeSet(oTree As SaffronTree) As clsRangeSet
'    Dim oRange As clsRange
'    Dim oSubTree As SaffronTree
'
'    Set DoRangeSet = New clsRangeSet
'
'    For Each oSubTree In oTree.SubTree
'        Set oRange = New clsRange
'        oRange.Starting = Val(oSubTree(1).Text)
'        oRange.Ending = Val(oSubTree(2).Text)
'        DoRangeSet.AddRange oRange
'    Next
'End Function
'
'Private Sub DoForLoop(oTree As SaffronTree, ByVal oContext As clsContext)
'    Dim oIndexer As clsVar
'    Dim oIndexVar As clsVar
'    Dim oSubContext As New clsContext
'    Dim oForLoop As New clsForLoop
'    Dim oStatement As clsIStatement
'    Dim oExtent As clsExtent
'
'    Set oExtent = DoExtent(oTree(1), oContext)
'
'    Set oIndexVar = New clsVar
'    oIndexVar.Identifier = oTree(2).Text
'    oSubContext.moVars.AddVar oIndexVar
'
'    Set oForLoop.moContext = oSubContext
'    Set oForLoop.moIndexer = oExtent
'    Set oForLoop.moIndexerVar = oIndexVar
'
''    Dim oIterator As New clsIterator
''    Set oIterator.Member = oIndexer.Extent.Members(0)
''
''    Do
''        Debug.Print oIterator.mlRealValue
''    Loop Until Not oIterator.NextValue
'End Sub
'
