VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenerateIntermediateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moVariables As New clsNode
Private mlVariablePtr As Long
Private moBlocks As New clsNode
Private moICode As New clsNode ' Intermediate Code
Private moACode As New clsNode ' Assembly Code
Private mlBlockCounter As Long

Private Enum ElementTypes
    Number
    VariableName
    FunctionName
End Enum

' Intermediate Code
Private Enum IntermediateCodeParameters
    icpDirection
    icpMode
    icpOperation
    icpValue
End Enum

Private Enum Direction
    diNone
    diRead
    diWrite
    diReadDouble
    diWriteDouble
End Enum

Private Enum Modes
    mdNone
    mdConstant
    mdAddress
    mdStack
End Enum

Private Enum OperatorTypes
    opCopy
    opAdd
    opSub
    opAnd
    opOr
    opXor
    opReturn
    opLabel
    opVariable
    opSubroutine
    opJump
    opCompare
    opStackZ
    opJumpZero
End Enum

' Assembly
Private Enum AssemblyCodeParameters
    acpInstruction
    acpOperand
    acpMode
End Enum

Private Enum AddressModes
    amImplied
    amImmediate
    amAddress
    amIndexedX
    amIndexedY
    amIndirectX
    amIndirectY
    amIndirect
    amLabel
    amIgnore
    amMatch
End Enum

Private Enum OptimiseRemovalTypes
    ortBoth
    ortFirst
    ortSecond
End Enum

Public Sub CompileAssembly(oProgramBlock As Block)
    
'    TestOptimisation
'    Exit Sub

    CompileVariables oProgramBlock, , oProgramBlock.msName
    
    AddIC diNone, mdConstant, opLabel, &H200
    CompileSubBlockExpressions oProgramBlock
    AddIC diNone, mdAddress, opLabel, "stack"
    
    OptimiseIntermediateCode
    
    WriteAssemblyVariablesCode
    AddAC "LDX", "00h", amImmediate
    WriteAssemblyCode
    AddAC "FFFCh DW 0200h", opLabel, amImplied
    
    OptimiseAssemblyCode
    
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    
    'Set oTS = oFSO.CreateTextFile(App.Path & "\assembly.txt", True)
    Set oTS = oFSO.CreateTextFile("C:\Computer\Programming\Emulation\6502Em\assembly.txt", True)
    oTS.Write WriteAssemblyText
    
End Sub

Private Sub CompileVariables(ByVal oVariableBlock As Block, Optional ByVal lVariablePtr As Long = 0, Optional ByVal sBlockName As String)
    Dim lIndex As Long
    Dim oVariable As clsNode
    Dim oBlock As Block
    
    For lIndex = 0 To oVariableBlock.moVariables.Count - 1
        Set oVariable = oVariableBlock.moVariables.ItemPhysical(lIndex)
        moVariables.AddNew(, sBlockName & "_" & oVariable.TextKey).Value = Array(lVariablePtr, oVariable.Value)
        lVariablePtr = lVariablePtr + oVariable.Value
    Next
    
    For lIndex = 0 To oVariableBlock.moSubBlocks.Count - 1
        Set oBlock = oVariableBlock.moSubBlocks.ItemPhysical(lIndex).Value
        CompileVariables oBlock, lVariablePtr, sBlockName & "_" & oVariableBlock.moSubBlocks.ItemPhysical(lIndex).TextKey
    Next
End Sub

Private Sub CompileSubBlockExpressions(oBlock As Block)
    Dim lIndex As Long
    Dim oSubBlock As Block
    
    AddIC diNone, mdAddress, opLabel, oBlock.msName
    
    CompileExpression oBlock

    For lIndex = 0 To oBlock.moSubBlocks.Count - 1
        Set oSubBlock = oBlock.moSubBlocks.ItemPhysical(lIndex).Value
        CompileSubBlockExpressions oSubBlock
    Next
End Sub

Private Sub AddIC(ByVal diDirection As Direction, ByVal otMode As Modes, ByVal opOperation As OperatorTypes, Optional ByVal sValue As String)
    moICode.AddNew().Value = Array(diDirection, otMode, opOperation, sValue)
End Sub

Private Sub RemoveIntermediateCode()
   moICode.RemovePhysical moICode.Count - 1
End Sub

Private Function CompileExpression(oBlock As Block, Optional bInline As Boolean)
    Dim oExpressionElement As clsNode
    Dim lIndex As Long
    Dim vVariable As Variant
    Dim lStackType1 As Long
    Dim lStackType2 As Long
    Dim lResult  As Long
    Dim oExpression As clsNode
    Dim oParentBlock As Block
    
    Set oExpression = oBlock.moIntermediateCode
    
    For lIndex = 0 To oExpression.Count - 1
        Set oExpressionElement = oExpression.ItemPhysical(lIndex)
        If TypeOf oExpressionElement.Value Is Block Then
            Select Case oExpressionElement.Value.btBlockType
                Case btBlock
                    CompileExpression oExpressionElement.Value, True
                Case btLoop
                    AddIC diNone, mdAddress, opLabel, oExpressionElement.Value.msName
                    CompileExpression oExpressionElement.Value, True
                    AddIC diNone, mdNone, opJump, oExpressionElement.Value.msName
                    AddIC diNone, mdAddress, opLabel, oExpressionElement.Value.msName & "_exit"
                Case btConditional
                    AddIC diRead, mdStack, opCopy
                    AddIC diNone, mdAddress, opJumpZero, oExpressionElement.Value.msName & "_exit"
                    CompileExpression oExpressionElement.Value, True
                    AddIC diNone, mdAddress, opLabel, oExpressionElement.Value.msName & "_exit"
            End Select
        ElseIf IsNumeric(oExpressionElement.Value) Then
            AddIC diRead, mdConstant, opCopy, Val(oExpressionElement.Value)
            AddIC diWrite, mdStack, opCopy
        Else
            vVariable = FindVariable(oExpressionElement.Value, oBlock)
            If Not IsEmpty(vVariable) Then
                If vVariable(1) = 1 Then ' short
                    AddIC diRead, mdAddress, opCopy, vVariable(0)
                    AddIC diWrite, mdStack, opCopy
                ElseIf vVariable(1) = 2 Then ' long
                    AddIC diReadDouble, mdAddress, opCopy, vVariable(0)
                    AddIC diWriteDouble, mdStack, opCopy
                End If
            Else
                ' Built In functions
                Select Case LCase$(oExpressionElement.Value)
                    Case "+"
                        AddIC diRead, mdStack, opCopy
                        AddIC diRead, mdStack, opAdd
                        AddIC diWrite, mdStack, opCopy
                    Case "-"
                        AddIC diRead, mdStack, opCopy
                        AddIC diRead, mdStack, opSub
                        AddIC diWrite, mdStack, opCopy
                    Case ":="
                        RemoveIntermediateCode
                        RemoveIntermediateCode
                        AddIC diRead, mdStack, opCopy
                        vVariable = FindVariable(oExpression.ItemPhysical(lIndex - 1).Value, oBlock)
                        AddIC diWrite, mdAddress, opCopy, vVariable(0)
                    Case "exit"
                        Set oParentBlock = oBlock
                        While oParentBlock.btBlockType <> btLoop And oParentBlock.btBlockType <> btFunction
                            Set oParentBlock = oBlock.moParentBlock
                        Wend
                        AddIC diNone, mdNone, opJump, oParentBlock.msName & "_exit"
                    Case "=" ' equality
                        AddIC diRead, mdStack, opCopy
                        AddIC diRead, mdStack, opCompare
                        AddIC diWrite, mdStack, opStackZ
                    Case Else ' Try user defined functions
                        AddIC diNone, mdNone, opSubroutine, oExpressionElement.Value
                End Select
            End If
        End If
    Next
    
    If Not bInline Then
        AddIC diNone, mdNone, opReturn
    End If
End Function

Private Function FindVariable(sName As String, oBlock As Block) As Variant
    Dim lIndex As Long
    Dim oParentBlock As Block
    Dim sQualifiedName As String
    Dim vVarInfo As Variant
    
    FindVariable = Empty
    
    ' Find fully qualified variable
    Set oParentBlock = oBlock
    While Not oParentBlock Is Nothing
        sQualifiedName = oParentBlock.msName & "_" & sQualifiedName
        Set oParentBlock = oParentBlock.moParentBlock
    Wend
    
    If Not moVariables.Keys.Word(sQualifiedName & sName) Is Nothing Then
        vVarInfo = moVariables.Keys.Word(sQualifiedName & sName).Value.Value
        FindVariable = Array(sQualifiedName & sName, vVarInfo(1))
        Exit Function
    End If
    
    ' Else look at parent scope
    If Not oBlock.moParentBlock Is Nothing Then
        FindVariable = FindVariable(sName, oBlock.moParentBlock)
    End If
End Function


Public Function HexNum(ByVal lNumber As Long, ByVal iPlaces As Integer) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
End Function


Private Sub WriteAssemblyVariablesCode()
    Dim lIndex As Long
    Dim oVariable As clsNode
    
    For lIndex = 0 To moVariables.Count - 1
        Set oVariable = moVariables.ItemPhysical(lIndex)
        AddAC HexNum(oVariable.Value(0), 4) & "h " & oVariable.TextKey & Choose(oVariable.Value(1), " DB 00h", " DW 0000h"), "", amImplied
    Next
End Sub


' Returns code length
Private Function WriteAssemblyCode() As Long
    Dim lInstruction As Long
    Dim vInstruction As Variant
    Dim sCode As String
    Dim sType As String
    
    For lInstruction = 0 To moICode.Count - 1
        vInstruction = moICode.ItemPhysical(lInstruction).Value
        
        Select Case vInstruction(0)
            Case diNone
                Select Case vInstruction(2)
                    Case opReturn
                        AddAC "RTS", , amImplied
                    Case opLabel
                        Select Case vInstruction(1)
                            Case mdAddress
                                AddAC vInstruction(3), , amLabel
                            Case mdConstant
                                AddAC HexNum(vInstruction(3), 4) & "h", , amLabel
                        End Select
                    Case opSubroutine
                        AddAC "JSR", vInstruction(3), amAddress
                    Case opJump
                        AddAC "JMP", vInstruction(3), amAddress
                    Case opJumpZero
                        AddAC "BEQ", vInstruction(3), amAddress
                End Select
            Case diRead
                Select Case vInstruction(2)
                    Case opCopy
                        sType = "LDA"
                    Case opAdd
                        AddAC "CLC", "", amImplied
                        sType = "ADC"
                    Case opSub
                        AddAC "SEC", "", amImplied
                        sType = "SBC"
                    Case opAnd
                        sType = "AND"
                    Case opOr
                        sType = "ORA"
                    Case opXor
                        sType = "EOR"
                    Case opCompare
                        sType = "CMP"
                End Select
                Select Case vInstruction(1)
                    Case mdConstant
                        AddAC sType, HexNum(vInstruction(3), 2) & "h", amImmediate
                    Case mdAddress
                        AddAC sType, vInstruction(3), amAddress
                    Case mdStack
                        AddAC "DEX"
                        AddAC sType, "stack", amIndexedX
                End Select
            Case diWrite
                Select Case vInstruction(1)
                    Case mdAddress
                        AddAC "STA", vInstruction(3), amAddress
                    Case mdStack
                        Select Case vInstruction(2)
                            Case opCopy
                                AddAC "STA", "stack", amIndexedX
                                AddAC "INX"
                            Case opStackZ
                                AddAC "PHP"
                                AddAC "PLA"
                                AddAC "AND", "02h", amImmediate
                                AddAC "STA", "stack", amIndexedX
                                AddAC "INX"
                        End Select
                End Select
            Case diReadDouble
            Case diWriteDouble
        End Select
    Next
End Function

Private Sub AddAC(ByVal sInstruction As String, Optional ByVal sOperand As String, Optional ByVal amMode As AddressModes = amImplied)
    moACode.AddNew().Value = Array(sInstruction, sOperand, amMode)
End Sub

Private Function WriteAssemblyText() As String
    Dim lIndex As Long
    
    For lIndex = 0 To moACode.Count - 1
        WriteAssemblyText = WriteAssemblyText & moACode.ItemPhysical(lIndex).Value(acpInstruction)
        
        Select Case moACode.ItemPhysical(lIndex).Value(acpMode)
            Case amImmediate
                WriteAssemblyText = WriteAssemblyText & " #" & moACode.ItemPhysical(lIndex).Value(acpOperand)
            Case amAddress
                WriteAssemblyText = WriteAssemblyText & " " & moACode.ItemPhysical(lIndex).Value(acpOperand)
            Case amImplied
                WriteAssemblyText = WriteAssemblyText
            Case amIndexedX
                WriteAssemblyText = WriteAssemblyText & " " & moACode.ItemPhysical(lIndex).Value(acpOperand) & ",X"
            Case amIndexedY
                WriteAssemblyText = WriteAssemblyText & " " & moACode.ItemPhysical(lIndex).Value(acpOperand) & ",Y"
            Case amIndirect
                WriteAssemblyText = WriteAssemblyText & " (" & moACode.ItemPhysical(lIndex).Value(acpOperand) & ")"
            Case amIndirectX
                WriteAssemblyText = WriteAssemblyText & " (" & moACode.ItemPhysical(lIndex).Value(acpOperand) & ",X)"
            Case amIndirectY
                WriteAssemblyText = WriteAssemblyText & " (" & moACode.ItemPhysical(lIndex).Value(acpOperand) & "),Y"
            Case amLabel
            
        End Select
        
        WriteAssemblyText = WriteAssemblyText & vbCrLf
    Next
End Function

Private Sub OptimiseIntermediateCode()
    Dim lIndex As Long
    Dim vIns1 As Variant
    Dim vIns2 As Variant

    While lIndex <= (moICode.Count - 2)
        vIns1 = moICode.ItemPhysical(lIndex).Value
        vIns2 = moICode.ItemPhysical(lIndex + 1).Value
        
        ' Stack push/pull pair
        If vIns1(icpMode) = mdStack And vIns2(icpMode) = mdStack Then
            If vIns1(icpOperation) = opCopy And vIns2(icpOperation) = opCopy Then
                If vIns1(icpDirection) = diRead And vIns2(icpDirection) = diWrite Then
                    moICode.RemovePhysical lIndex
                    moICode.RemovePhysical lIndex
                    lIndex = lIndex - 2
                ElseIf vIns1(icpDirection) = diWrite And vIns2(icpDirection) = diRead Then
                    moICode.RemovePhysical lIndex
                    moICode.RemovePhysical lIndex
                    lIndex = lIndex - 2
                End If
            End If
        End If
        lIndex = lIndex + 1
    Wend
End Sub

Private Sub TestOptimisation()

    AddAC "STA", "stack", amIndexedX
    AddAC "BEQ", "stack", amAddress
    AddAC "INX", "", amImplied
    AddAC "LDA", "stack", amIndexedX
    AddAC "INX", "", amImplied
    
    OptimiseAssemblyCode
    Debug.Print WriteAssemblyText
End Sub

Private Sub OptimiseAssemblyCode()
    Dim vFlowExceptions As Variant
    Dim vExceptionsINX As Variant
    Dim vExceptionsSTA As Variant
    Dim lIndex As Long
    
    vFlowExceptions = Array()
    ExceptionsControlFlow vFlowExceptions
    
    vExceptionsINX = vFlowExceptions
    ExceptionsDependsOnX vExceptionsINX

    vExceptionsSTA = vFlowExceptions
    ExceptionsChangesA vExceptionsSTA
    
    Do While lIndex < moACode.Count
        EliminatePair lIndex, Array("INX", "", amImplied), Array("DEX", "", amImplied), ortBoth, vExceptionsINX
        EliminatePair lIndex, Array("STA", "amIgnore", amAddress), Array("LDA", "amMatch", amMatch), ortSecond, vExceptionsSTA
        
        lIndex = lIndex + 1
    Loop
End Sub

Private Sub ExceptionsControlFlow(vExceptions)
    ' Control flow
    AppendArray vExceptions, Array("RTS", "", amImplied)
    AppendArray vExceptions, Array("JMP", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("JSR", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BEQ", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BNE", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BCC", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BCS", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BVC", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BVS", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BPL", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("BMI", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("amIgnore", "amIgnore", amLabel)
End Sub

Private Function ExceptionsChangesA(vExceptions As Variant)
    AppendArray vExceptions, Array("LDA", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("ADC", "amIgnore", amIgnore)
    AppendArray vExceptions, Array("SBC", "amIgnore", amIgnore)
End Function

Private Function ExceptionsDependsOnX(vExceptions As Variant)
    AppendArray vExceptions, Array("amIgnore", "amIgnore", amIndexedX)
    AppendArray vExceptions, Array("amIgnore", "amIgnore", amIndirectX)
End Function

Private Sub AppendArray(vArray As Variant, vItem As Variant)
    ReDim Preserve vArray(UBound(vArray) + 1)
    vArray(UBound(vArray)) = vItem
End Sub

Private Sub EliminatePair(ByVal lStartLine As Long, ByVal vIns1 As Variant, ByVal vIns2 As Variant, ByVal ortType As OptimiseRemovalTypes, ByVal vExceptions As Variant)
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim lIndex3 As Long
    Dim vInsLine1 As Variant
    Dim vInsLine2 As Variant
    Dim vException As Variant
    Dim vExcepted As Variant
    
    vInsLine1 = moACode.ItemPhysical(lStartLine).Value
    If Not InstructionMatch(vInsLine1, vIns1) Then
        Exit Sub
    End If

    lIndex2 = lStartLine + 1
    
    Do While lIndex2 < moACode.Count
        vInsLine2 = moACode.ItemPhysical(lIndex2).Value

        If InstructionMatch(vInsLine2, vIns2, vInsLine1) Then
            Exit Do
        End If
        
        For Each vException In vExceptions
            If InstructionMatch(vInsLine2, vException) Then
                Exit Sub
            End If
        Next
        
        lIndex2 = lIndex2 + 1
    Loop
    
    If lIndex2 < moACode.Count Then
        If ortType = ortBoth Or ortType = ortSecond Then
            moACode.RemovePhysical lIndex2
        End If
        
        If ortType = ortBoth Or ortType = ortFirst Then
            moACode.RemovePhysical lStartLine
        End If
    End If
End Sub

Private Function InstructionMatch(vIns1 As Variant, vIns2 As Variant, Optional vInsMatch As Variant) As Boolean
    Dim vIgnore1 As Variant
    Dim vIgnore2 As Variant
    Dim vIgnoreCombined As Variant
    Dim vMatched As Variant
    
    vIgnore1 = InstructionIgnore(vIns1)
    vIgnore2 = InstructionIgnore(vIns2)
    vIgnoreCombined = Array(vIgnore1(acpInstruction) Or vIgnore2(acpInstruction), vIgnore1(acpOperand) Or vIgnore2(acpOperand), vIgnore1(acpMode) Or vIgnore2(acpMode))
    
    vMatched = Array(True, True, True)
    If Not vIgnoreCombined(acpInstruction) Then
        If vIns2(acpInstruction) = "amMatch" Then
            If vIns1(acpInstruction) <> vInsMatch(acpInstruction) Then
                vMatched(acpInstruction) = False
            End If
        Else
            If vIns1(acpInstruction) <> vIns2(acpInstruction) Then
                vMatched(acpInstruction) = False
            End If
        End If
    End If
    
    If Not vIgnoreCombined(acpMode) Then
        If vIns2(acpMode) = amMatch Then
            If vIns1(acpMode) <> vInsMatch(acpMode) Then
                vMatched(acpMode) = False
            End If
        Else
            If vIns1(acpMode) <> vIns2(acpMode) Then
                vMatched(acpMode) = False
            End If
        End If
    End If
    
    If Not vIgnoreCombined(acpOperand) Then
        If vIns2(acpOperand) = "amMatch" Then
            If vIns1(acpOperand) <> vInsMatch(acpOperand) Then
                vMatched(acpOperand) = False
            End If
        Else
            If vIns1(acpOperand) <> vIns2(acpOperand) Then
                vMatched(acpOperand) = False
            End If
        End If
    End If
    
    InstructionMatch = vMatched(acpInstruction) And vMatched(acpMode) And vMatched(acpOperand)
End Function

Private Function InstructionIgnore(vIns As Variant) As Variant
    Dim vExcepted As Variant
    
    vExcepted = Array(False, False, False)
    
    If vIns(acpInstruction) = "amIgnore" Then
        vExcepted(acpInstruction) = True
    End If
    
    If vIns(acpMode) = amIgnore Then
        vExcepted(acpMode) = True
    End If

    If vIns(acpOperand) = "amIgnore" Then
        vExcepted(acpOperand) = True
    End If
    
    InstructionIgnore = vExcepted
End Function
