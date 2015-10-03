Attribute VB_Name = "Intermediate"
Option Explicit

Private lCurrentObjectAddress As Long
Private lBaseAddress As Long
Private mvOutput As Variant
  
Function Compile(oContext As clsContext) As String
    Dim vOutput As Variant
    Dim oIntermediate As clsIntermediate
    Dim lOperand1Address As Long
    Dim lOperand2Address As Long
    Dim lOperand3Address As Long
    Dim lIndex As Long
    
    lCurrentObjectAddress = &H70
    lBaseAddress = &HC00
    mvOutput = Array()
    
    Variables oContext
          
    AddLine HexNum(lBaseAddress)
    
    For lIndex = 0 To oContext.moIntermediates.Count - 1
        Set oIntermediate = oContext.moIntermediates.Intermediates(lIndex)
        Select Case oIntermediate.Operator
            Case opCopy
                OperatorCopy oIntermediate
            Case opAdd
                OperatorAdd oIntermediate
            Case opSub
                OperatorSubtract oIntermediate
            Case opMultiply
                OperatorMultiply oIntermediate
            Case opDivide
                OperatorDivide oIntermediate
            Case opModulus
                OperatorModulus oIntermediate
        End Select
    Next
    
    AddLine "RTS"
     
    Compile = Join(mvOutput, vbCrLf)
End Function

Private Function NextObjectAddress(lByteSize As Long)
    NextObjectAddress = lCurrentObjectAddress
    lCurrentObjectAddress = lCurrentObjectAddress + lByteSize
End Function


Private Sub AddLine(sLine As String)
    ReDim Preserve mvOutput(UBound(mvOutput) + 1)
    mvOutput(UBound(mvOutput)) = sLine
End Sub

Public Function HexNum(ByVal lNumber As Long, Optional ByVal iPlaces As Integer = 4) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
    HexNum = HexNum & "h"
End Function

Private Function Label(sIdentify As String) As String
    Dim lIndex As Long
    
    For lIndex = 1 To 5
        Label = Label & Chr$(Rnd() * 26 + 65)
    Next
    Label = sIdentify & "_" & Label
End Function

Private Function Var(sIdentify As String) As String
    Var = "var_" & sIdentify
End Function

Private Function Variables(oContext As clsContext)
    Dim lIndex As Long
    Dim lIndex2
    Dim oObject As clsObject
    Dim sLine As String
    Dim lSize As Long
    Dim lValue As Long
    
    For lIndex = 0 To oContext.moObjects.Count - 1
        Set oObject = oContext.moObjects.Objects(lIndex)
        sLine = HexNum(oObject.MyAddress(NextObjectAddress(oObject.UnitClass.Size)), 4) & " var_" & oObject.Identifier
        If oObject.IsConstant Then
            lValue = oObject.UnitClass.Range.Starting
            For lIndex2 = 1 To oObject.UnitClass.Size
                sLine = sLine & " DB " & HexNum(lValue And 255, 2)
                lValue = lValue \ 256
            Next
        End If
        AddLine sLine
    Next
End Function

Private Function Largest(ByVal lValue1 As Long, ByVal lValue2 As Long) As Long
    If lValue1 > lValue2 Then
        Largest = lValue1
    Else
        Largest = lValue2
    End If
End Function

Private Function Offset(ByVal sIdentifier, ByVal lOffset As Long)
    If lOffset > 0 Then
        Offset = sIdentifier & "+" & lOffset & "d"
    Else
        Offset = sIdentifier
    End If
End Function

Private Function OperatorCopy(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOffset As Long
    
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    lOperand1Size = oIntermediate.Operand1.UnitClass.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Size
    
    For lOffset = 1 To Largest(lOperand1Size, lOperand2Size)
        If lOffset <= lOperand2Size Then
            AddLine "LDA " & Offset(sOperand2, lOffset - 1)
        Else
            AddLine "LDA #00h"
        End If
        If lOffset <= lOperand1Size Then
            AddLine "STA " & Offset(sOperand1, lOffset - 1)
        End If
    Next
End Function

Private Function OperatorAdd(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOffset As Long
    
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    sOperand3 = Var(oIntermediate.Operand3.Identifier)

    lOperand1Size = oIntermediate.Operand1.UnitClass.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Size
    
    For lOffset = 1 To Largest(lOperand1Size, lOperand2Size)
        If lOffset <= lOperand1Size Then
            AddLine "LDA " & Offset(sOperand1, lOffset - 1)
        Else
            AddLine "LDA #00h"
        End If
        If lOffset <= lOperand2Size Then
            If lOffset = 1 Then
                AddLine "CLC"
            End If
            AddLine "ADC " & Offset(sOperand2, lOffset - 1)
        Else
            AddLine "ADC #00h"
        End If
        If lOffset <= lOperand3Size Then
            AddLine "STA " & Offset(sOperand3, lOffset - 1)
        End If
    Next
End Function

Private Function OperatorSubtract(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOffset As Long
    
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    sOperand3 = Var(oIntermediate.Operand3.Identifier)

    lOperand1Size = oIntermediate.Operand1.UnitClass.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Size
    
    For lOffset = 1 To Largest(lOperand1Size, lOperand2Size)
        If lOffset <= lOperand1Size Then
            AddLine "LDA " & Offset(sOperand1, lOffset - 1)
        Else
            AddLine "LDA #00h"
        End If
        If lOffset <= lOperand2Size Then
            If lOffset = 1 Then
                AddLine "SEC"
            End If
            AddLine "SBC " & Offset(sOperand2, lOffset - 1)
        Else
            AddLine "SBC #00h"
        End If
        If lOffset <= lOperand3Size Then
            AddLine "STA " & Offset(sOperand3, lOffset - 1)
        End If
    Next
End Function

Private Function OperatorMultiply(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    
    Dim sLabel1 As String
    Dim sLabel2 As String
        
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    sOperand3 = Var(oIntermediate.Operand3.Identifier)
    
    sLabel1 = Label("mult_next")
    sLabel2 = Label("mult_noadd")
    
    AddLine "LDA " & sOperand1
    AddLine "STA " & sOperand3
    AddLine "LDX #08h"
    AddLine "LDA #00h"
    AddLine "CLC"
    AddLine sLabel1
    AddLine "BCC " & sLabel2
    AddLine "CLC"
    AddLine "ADC " & sOperand2
    AddLine sLabel2
    AddLine "ROR A"
    AddLine "ROR " & sOperand3
    AddLine "DEX"
    AddLine "BPL " & sLabel1
    'AddLine "STA 71h"

End Function

Private Function OperatorDivide(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim sLabel1 As String
    Dim sLabel2 As String
        
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    sOperand3 = Var(oIntermediate.Operand3.Identifier)
    
    sLabel1 = Label("divide_next")
    sLabel2 = Label("divide_nosub")
    
    AddLine "LDA " & sOperand1
    AddLine "STA " & sOperand3
    AddLine "LDX #8d"
    AddLine "LDA " & sOperand2
    AddLine sLabel1
    AddLine "CMP " & sOperand2
    AddLine "BCC " & sLabel2
    AddLine "SBC " & sOperand2
    AddLine sLabel2
    'AddLine "ROL 70h"
    AddLine "ROL " & sOperand3
    AddLine "ROL A"
    AddLine "DEX"
    AddLine "BPL " & sLabel1

End Function

Private Function OperatorModulus(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim sLabel1 As String
    Dim sLabel2 As String
    
    sOperand1 = Var(oIntermediate.Operand1.Identifier)
    sOperand2 = Var(oIntermediate.Operand2.Identifier)
    sOperand3 = Var(oIntermediate.Operand3.Identifier)
    
    sLabel1 = Label("modulus_next")
    sLabel2 = Label("modulus_nosub")
    
    AddLine "LDA " & sOperand1
    AddLine "STA " & sOperand3
    AddLine "LDX #8d"
    AddLine "LDA " & sOperand2
    AddLine sLabel1
    AddLine "CMP " & sOperand2
    AddLine "BCC " & sLabel2
    AddLine "SBC " & sOperand2
    AddLine sLabel2
    'AddLine "ROL 70h"
    AddLine "ROL " & sOperand3
    AddLine "ROL A"
    AddLine "DEX"
    AddLine "BPL " & sLabel1
    AddLine "ROR A"
    AddLine "STA " & sOperand3
    
End Function

