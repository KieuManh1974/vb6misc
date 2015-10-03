Attribute VB_Name = "Program"
Option Explicit

'Private ProgramAnalyse As New Analyse
'Private ProgramCompose As New Compose
'Private ProgramGenerate As New GenerateIntermediateCode

Private Types As New Collection
Private Vars As New Collection

Public Function LoadProgram()
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
End Function

Private Sub DoStatements(oTree As SaffronTree, oContext As clsContext)
    Dim oStatement As SaffronTree
    
    For Each oStatement In oTree.SubTree
        Select Case oStatement.Index
            Case 1 ' Declaration
                DoDeclaration oStatement(1), oContext
            Case 2 ' For loop
                DoForLoop oStatement(1), oContext
        End Select
    Next
End Sub

Private Sub DoDeclaration(oTree As SaffronTree, oContext As clsContext)
    Select Case oTree(1).Index
        Case 1 ' rec
            DoRec oTree, oContext
        Case 2 ' str
            DoStr oTree, oContext
        Case 3 ' var
            DoVar oTree, oContext
    End Select
End Sub

Private Sub DoRec(oTree As SaffronTree, oContext As clsContext)
    Dim oRec As New clsRec
    
    oRec.Identifier = oTree(2).Text
    Set oRec.Extent = DoExtent(oTree(3), oContext)
    
    oContext.moRecs.AddRec oRec
End Sub

Private Sub DoStr(oTree As SaffronTree, oContext As clsContext)
    Dim oStr As New clsStr

    oStr.Identifier = oTree(2).Text
    Set oStr.Extent = DoExtent(oTree(3), oContext)
    
    oContext.moStrs.AddStr oStr
End Sub

Private Sub DoVar(oTree As SaffronTree, oContext As clsContext)
    Dim oVar As New clsVar
    
    oVar.Identifier = oTree(2).Text
    Set oVar.VarArray.Extents = DoExtent(oTree(3), oContext)
    
    oContext.moVars.AddVar oVar
End Sub

Private Function DoExtent(oTree As SaffronTree, oContext As clsContext) As clsExtents
    Dim oExtent As clsExtent
    Dim oExtents As New clsExtents
    Dim oSubTree As SaffronTree
    Dim oSet As Collection
    Dim oMember As clsIMember
    Dim lIndex As Long
    
    lIndex = 1
    For Each oSubTree In oTree.SubTree
        Select Case oSubTree.Index
            Case 1 ' index set
                Set oExtent = New clsExtent
                Set oSet = DoIndexSet(oSubTree)
                For Each oMember In oSet
                    oExtent.AddMember oMember
                Next
                oExtents.AddExtent oExtent
                
            Case 2 ' identifier ' first must be a rec, others must be str
                Select Case lIndex
                    Case 1
                        Set oExtent = New clsExtent
                        oExtent.AddMember oContext.moRecs.GetByIdentifier(oSubTree.Text)
                        oExtents.AddExtent oExtent
                    Case 2
                        Set oExtent = New clsExtent
                        oExtent.AddMember oContext.moStrs.GetByIdentifier(oSubTree.Text)
                        oExtents.AddExtent oExtent
                End Select
        End Select
        lIndex = lIndex + 1
    Next
    
    Set DoExtent = oExtents
End Function

Private Function DoIndexSet(oTree As SaffronTree) As Collection
    Dim oSubTree As SaffronTree
    Dim oRangeSet As clsRangeSet
    Dim oSymbolSet As clsSymbolSet
    Dim oIdentifierSet As clsIdentifierSet
    Dim sIdentifier As String
    Dim oIdentifier As SaffronTree
    Dim oMember As clsIMember
    
    Set DoIndexSet = New Collection
    
    For Each oSubTree In oTree(1)(1).SubTree
        sIdentifier = oSubTree(1).Text
        
        Select Case oSubTree(2).Index
            Case 1 ' symbol set
                Set oSymbolSet = New clsSymbolSet
                oSymbolSet.SymbolSet = oSubTree(2).Text
                Set oMember = oSymbolSet
                oMember.Identifier = sIdentifier
            Case 2 ' identifier set
                Set oIdentifierSet = New clsIdentifierSet
                For Each oIdentifier In oSubTree(2)(1).SubTree
                    oIdentifierSet.AddIdentifier oIdentifier.Text
                Next
                Set oMember = oIdentifierSet
                oMember.Identifier = sIdentifier
            Case 3 ' range set
                Set oRangeSet = DoRangeSet(oSubTree(2)(1))
                Set oMember = oRangeSet
                oMember.Identifier = sIdentifier
        End Select
        
        DoIndexSet.Add oMember
    Next
End Function

Private Function DoRangeSet(oTree As SaffronTree) As clsRangeSet
    Dim oRange As clsRange
    Dim oSubTree As SaffronTree
    
    Set DoRangeSet = New clsRangeSet
    
    For Each oSubTree In oTree.SubTree
        Set oRange = New clsRange
        oRange.Starting = Val(oSubTree(1).Text)
        oRange.Ending = Val(oSubTree(2).Text)
        DoRangeSet.AddRange oRange
    Next
End Function

Private Sub DoForLoop(oTree As SaffronTree, ByVal oContext As clsContext)
    Dim oIndexer As clsVar
    Dim oIndexVar As clsVar
    Dim oSubContext As New clsContext
    Dim oForLoop As New clsForLoop
    Dim oStatement As clsIStatement
    Dim oExtent As clsExtent
    
    Set oExtent = DoExtent(oTree(1), oContext)
    
    Set oIndexVar = New clsVar
    oIndexVar.Identifier = oTree(2).Text
    oSubContext.moVars.AddVar oIndexVar
    
    Set oForLoop.moContext = oSubContext
    Set oForLoop.moIndexer = oExtent
    Set oForLoop.moIndexerVar = oIndexVar

'    Dim oIterator As New clsIterator
'    Set oIterator.Member = oIndexer.Extent.Members(0)
'
'    Do
'        Debug.Print oIterator.mlRealValue
'    Loop Until Not oIterator.NextValue
End Sub

