Attribute VB_Name = "modCompile"
Option Explicit

Private mlAdr As Long
Private mlPtr As Long

Public Sub Compile()
    Dim sProgram As String
    Dim lAddress As Long
    Dim lMemory As Long
    Dim lBit As Long
    Dim lValue As Long
    Dim lIndex As Long
    
'    sProgram = NandOp(131080, 131072)
'    sProgram = sProgram & Jump(0)
    
'    'sProgram = Nand(6 + 16, 16)
    For lIndex = 0 To 7
        sProgram = sProgram & FullAddOp(131080 + lIndex, 131072 + lIndex)
    Next
    sProgram = sProgram & NandOp(Temp(), Zero())
    sProgram = sProgram & NotOp(Temp())
    sProgram = sProgram & Jump(0)
    
    'Debug.Print sProgram
    For lIndex = 1 To Len(sProgram)
        lBit = lAddress And 7&
        lMemory = lAddress \ 8&
        Select Case Mid$(sProgram, lIndex, 1)
            Case "L"
                lValue = 0
            Case "O"
                lValue = 1
            Case "R"
                lValue = 2
            Case "J"
                lValue = 3
        End Select
        If lBit < 7 Then
            Memory(lMemory) = Memory(lMemory) And (3 * (myLookup(lBit)) Xor &HFF) Or lValue * myLookup(lBit)
        Else
            Memory(lMemory) = Memory(lMemory) And &H7F Or 128 * (lValue And 1)
            Memory(lMemory + 1) = Memory(lMemory + 1) And &HFE Or (lValue \ 2)
        End If
        lAddress = lAddress + 2
    Next
    Debug.Print lIndex  ' 103 instructions
End Sub

Private Function Zero() As Long
    Zero = 65535 * 8 + 7
End Function

Public Function Temp(Optional lIndex As Long = 0) As Long
    Temp = 65535 * 8 + 6 - lIndex
End Function

Private Function NandOp(lAddress2 As Long, lAddress1 As Long) As String
    NandOp = Rep(lAddress1) & "O" & Rep(lAddress2) & "O"
End Function

Private Function Nand2Op(lAddress3 As Long, lAddress2 As Long, lAddress1 As Long) As String
    Nand2Op = CopyOp(lAddress3, lAddress2) & NandOp(lAddress3, lAddress1)
End Function

Private Function XorOp(lAddress2 As Long, lAddress1 As Long) As String
    XorOp = Nand2Op(Temp(), lAddress2, lAddress1) & NandOp(lAddress2, Temp()) & NandOp(Temp(), lAddress1) & NandOp(lAddress2, Temp())
End Function

Private Function NotOp(lAddress As Long) As String
    NotOp = Rep(lAddress) & "O" & Rep(lAddress) & "O"
End Function

Private Function AndOp(lAddress2 As Long, lAddress1 As Long) As String
    AndOp = NandOp(lAddress2, lAddress1) & NotOp(lAddress2)
End Function

Private Function OrOp(lAddress2 As Long, lAddress1 As Long) As String
    OrOp = NandOp(Temp(), Zero()) & NandOp(Temp(), lAddress1) & NotOp(lAddress2) & NandOp(lAddress2, Temp())
End Function

Private Function NorOp(lAddress2 As Long, lAddress1 As Long) As String
    NorOp = NandOp(Temp(), Zero()) & NandOp(Temp(), lAddress1) & NotOp(lAddress2) & NandOp(lAddress2, Temp()) & NotOp(lAddress2)
End Function

Private Function FullAddOp(lAddress2 As Long, lAddress1 As Long) As String
    Dim A As Long
    Dim B As Long
    Dim C As Long
    Dim D As Long
    Dim E As Long
    Dim F As Long
    Dim T As Long
    
    A = lAddress1
    B = lAddress2
    C = Temp(1)
    D = Temp(2)
    E = Temp(3)
    F = Temp(4)
    T = Temp()
    
    
    FullAddOp = Nand2Op(C, B, A)
    FullAddOp = FullAddOp & Nand2Op(D, C, A)
    FullAddOp = FullAddOp & NandOp(B, C)
    FullAddOp = FullAddOp & NandOp(B, D)
    FullAddOp = FullAddOp & Nand2Op(E, B, T)
    FullAddOp = FullAddOp & NandOp(B, E)
    FullAddOp = FullAddOp & Nand2Op(F, T, E)
    FullAddOp = FullAddOp & NandOp(B, F)
    FullAddOp = FullAddOp & Nand2Op(T, E, C)
End Function

Private Function Jump(lAddress As Long) As String
    Jump = Rep(lAddress) & ResetPtr() & "J"
End Function

Private Function ResetPtr() As String
    ResetPtr = String$(mlPtr, "L")
End Function

Private Function CopyOp(lAddress2 As Long, lAddress1 As Long)
    CopyOp = NandOp(lAddress2, Zero()) & NandOp(lAddress2, lAddress1) & NandOp(lAddress2, lAddress2)
End Function

Private Function Rep(ByVal lDecimal As Long) As String
    Dim lIndex As Long
    Dim lSize As Long
    Dim lDigits() As Long
    Dim lThisDecimal As Long
    Dim lFirstOne As Long
    Dim lLastRPos As Long
    
    lThisDecimal = mlAdr Xor lDecimal
    mlAdr = lDecimal
    
    lFirstOne = -1
    Do
        ReDim Preserve lDigits(lSize)
        lDigits(lSize) = lThisDecimal And 1
        If lDigits(lSize) = 1 And lFirstOne = -1 Then
            lFirstOne = lSize
        End If
        lThisDecimal = lThisDecimal \ 2
        lSize = lSize + 1
    Loop Until lThisDecimal = 0
    
    lSize = lSize - 1
    
    If lSize >= mlPtr Then
        For lIndex = mlPtr To lSize
            If lDigits(lIndex) = 0 Then
                Rep = Rep & "RLR"
            Else
                Rep = Rep & "R"
            End If
        Next
        If lFirstOne < mlPtr Then
            For lIndex = lSize To mlPtr Step -1
                Rep = Rep & "L"
            Next
            For lIndex = mlPtr - 1 To lFirstOne Step -1
                If lDigits(lIndex) = 0 Then
                    Rep = Rep & "L"
                Else
                    Rep = Rep & "LRL"
                End If
            Next
            lLastRPos = InStrRev(Rep, "R")
            mlPtr = lFirstOne + (Len(Rep) - lLastRPos)
            Rep = Left$(Rep, lLastRPos)
        Else
            mlPtr = lSize + 1
        End If
    Else
        If lFirstOne <> -1 Then
            For lIndex = mlPtr To lSize + 1 Step -1
                Rep = Rep & "L"
            Next
            For lIndex = lSize To lFirstOne Step -1
                If lDigits(lIndex) = 0 Then
                    Rep = Rep & "L"
                Else
                    Rep = Rep & "RLL"
                End If
            Next
            
            lLastRPos = InStrRev(Rep, "R")
            mlPtr = lFirstOne - 1 + (Len(Rep) - lLastRPos)
            Rep = Left$(Rep, lLastRPos)
        End If
        
    End If
    
End Function
