Attribute VB_Name = "Compiler"
Option Explicit

Private mlVariableLocation As Long
Private mlTempLocation As Long

Private Type SymbolEntry
    Identifier As String
    Location As Long
    Size As Long
    Negative As Boolean
    Value As Long
End Type

Private mseExpressionStack() As SymbolEntry
Private mlExpressionStackSize As Long

Private mseSymbolTable() As SymbolEntry
Private mlSymbolTableSize As Long

Public Sub InitialiseCompiler()
    'Set moVariables = New clsNode
    'Set moTypes = New clsNode
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
    
    mlVariableLocation = &HC00&
    mlTempLocation = &HC80&
    
    For Each oSubTree In oTree.SubTree
        Select Case oSubTree.Index
            Case 1 ' declaration
                Declaration oSubTree(1)
            Case 2 ' expression
                Expression0 oSubTree(1)
        End Select
    Next
End Function

Private Sub Declaration(oTree As SaffronTree)
    Dim lType As Long
    Dim sIdentifier As String
    Dim seSymbol As SymbolEntry
    
    Select Case oTree(1).Index
        Case 1 ' SHORT
            seSymbol.Size = 1
            seSymbol.Negative = True
        Case 2 ' USHORT
            seSymbol.Size = 1
        Case 3 ' INT
            seSymbol.Size = 2
            seSymbol.Negative = True
        Case 4 ' UINT
            seSymbol.Size = 2
    End Select
    seSymbol.Identifier = oTree(2).Text
    seSymbol.Location = NewVariableAddress(seSymbol.Size)
    AddSymbol seSymbol
End Sub

Private Function NewVariableAddress(ByVal lSize As Long)
    NewVariableAddress = mlVariableLocation
    mlVariableLocation = mlVariableLocation + lSize
End Function

Private Function NewTempAddress(ByVal lSize As Long)
    NewTempAddress = mlTempLocation
    mlTempLocation = mlTempLocation + lSize
End Function

Private Sub Expression0(oTree As SaffronTree)
    Dim oPart As SaffronTree
    Dim seVar As SymbolEntry
    Dim seTemp As SymbolEntry
    Dim sPreviousOperator As String
    
    For Each oPart In oTree.SubTree
        Select Case oPart.Index
            Case 0 ' operator
                sPreviousOperator = oPart.Text
            Case 1 ' unary
                Select Case oPart(1)(1).Text
                    Case "?"
                    Case "-"
                    Case "_"
                End Select
            Case 2 ' function
            Case 3 ' var
                seVar = GetSymbol(oPart.Text)
                If seVar.Location <> -1 Then
                    PushExpressionStack seVar
                End If
            Case 4 ' constant
                seTemp.Location = NewTempAddress(2)
                seTemp.Size = 2
                seTemp.Value = Val(oPart.Text)
                PushExpressionStack seTemp
            Case 5 ' bracket
        End Select
        
        Select Case oPart.Index
            Case 1, 2, 3, 4
                Select Case sPreviousOperator
                    Case "+"
                        seTemp.Location = NewTempAddress(2)
                        seTemp.Size = 2
                        PushExpressionStack seTemp
                        EmitAdd
                        seTemp = GetStack(-1)
                        PullExpressionStack
                        PullExpressionStack
                        PullExpressionStack
                        PushExpressionStack seTemp
                    Case "-"
                        seTemp.Location = NewTempAddress(2)
                        seTemp.Size = 2
                        PushExpressionStack seTemp
                        EmitSub
                        seTemp = GetStack(-1)
                        PullExpressionStack
                        PullExpressionStack
                        PullExpressionStack
                        PushExpressionStack seTemp
            End Select
        End Select
    Next
End Sub

Private Sub PushExpressionStack(seSymbol As SymbolEntry)
    ReDim Preserve mseExpressionStack(mlExpressionStackSize)
    mseExpressionStack(mlExpressionStackSize) = seSymbol
    mlExpressionStackSize = mlExpressionStackSize + 1
End Sub

Private Sub PullExpressionStack()
    mlExpressionStackSize = mlExpressionStackSize - 1
    If mlExpressionStackSize <> 0 Then
        ReDim Preserve mseExpressionStack(mlExpressionStackSize - 1)
    Else
        Erase mseExpressionStack
    End If
End Sub

Private Sub AddSymbol(seSymbol As SymbolEntry)
    ReDim Preserve mseSymbolTable(mlSymbolTableSize)
    mseSymbolTable(mlSymbolTableSize) = seSymbol
    mlSymbolTableSize = mlSymbolTableSize + 1
End Sub

Private Function GetSymbol(ByVal sIdentifier As String) As SymbolEntry
    Dim lIndex As Long
    
    For lIndex = 0 To mlSymbolTableSize - 1
        If mseSymbolTable(lIndex).Identifier = sIdentifier Then
            GetSymbol = mseSymbolTable(lIndex)
            Exit Function
        End If
    Next
    
    GetSymbol.Location = -1
End Function

Private Function GetStack(ByVal lPosition As Long) As SymbolEntry
    GetStack = mseExpressionStack(mlExpressionStackSize + lPosition)
End Function

Private Sub EmitAdd()
    Dim lAddress As Long
    Dim seOperand1 As SymbolEntry
    Dim seOperand2 As SymbolEntry
    Dim seResult As SymbolEntry
    
    seResult = GetStack(-1)
    seOperand2 = GetStack(-2)
    seOperand1 = GetStack(-3)
    
    Emit "CLC"
    Emit "LDA " & HexNum(seOperand1.Location, 4) & "h"
    Emit "ADC " & HexNum(seOperand2.Location, 4) & "h"
    Emit "STA " & HexNum(seResult.Location, 4) & "h"
    Emit "LDA " & HexNum(seOperand1.Location + 1, 4) & "h"
    Emit "ADC " & HexNum(seOperand2.Location + 1, 4) & "h"
    Emit "STA " & HexNum(seResult.Location + 1, 4) & "h"
End Sub

Private Sub EmitSub()
    Dim lAddress As Long
    Dim seOperand1 As SymbolEntry
    Dim seOperand2 As SymbolEntry
    Dim seResult As SymbolEntry
    
    seResult = GetStack(-1)
    seOperand2 = GetStack(-2)
    seOperand1 = GetStack(-3)
    
    Emit "SEC"
    Emit "LDA " & seOperand1.Location
    Emit "SBC " & seOperand2.Location
    Emit "STA " & seResult.Location
    Emit "LDA " & seOperand1.Location + 1
    Emit "SBC " & seOperand2.Location + 1
    Emit "STA " & seResult.Location + 1
End Sub

Private Sub Emit(ByVal sLine As String)
    Debug.Print sLine
End Sub

Public Function HexNum(ByVal lNumber As Long, ByVal iPlaces As Integer) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
End Function

