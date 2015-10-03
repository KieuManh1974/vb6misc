Attribute VB_Name = "modNano"
Option Explicit

Public Adr As Long
Public Toggle As Long
Public Ctr As Long

Public myLookup(31) As Long
Private myFirstOperand As Byte
Private mlOperandIndex As Long

Public Sub Initialise()
    Dim lMultiplier As Long
    Dim lBitIndex As Long
    
    lMultiplier = 1
    For lBitIndex = 0 To 7
        myLookup(lBitIndex) = lMultiplier
        lMultiplier = lMultiplier * 2
    Next
    
    Toggle = 1
End Sub

Public Sub Execute()
    Dim lMemoryAddress As Long
    Dim lBitAddress As Long
    Dim lMask As Long
    Dim lInstruction As Long
    Dim lBit As Long

    Dim lThisBitAddress As Long
    Dim lThisMemoryAddress As Long
    
    Dim lMemory As Long
    
    StartCounter
    
    lMemory = Memory(lMemoryAddress)
    Do
        Select Case lMemory And 3
            Case 0 ' L
                If Toggle > 1& Then
                    Toggle = Toggle \ 2&
                End If
                lBitAddress = lBitAddress + 2
                If lBitAddress = 8 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = Memory(lMemoryAddress)
                    lBitAddress = 0
                ElseIf lBitAddress = 7 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = lMemory Or Memory(lMemoryAddress) * 2
                ElseIf lBitAddress = 9 Then
                    lBitAddress = 1
                Else
                    lMemory = lMemory \ 4
                End If
'                Debug.Print "L";
            Case 2 ' R
                Adr = Adr Xor Toggle
                Toggle = Toggle * 2
                lBitAddress = lBitAddress + 2
                If lBitAddress = 8 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = Memory(lMemoryAddress)
                    lBitAddress = 0
                ElseIf lBitAddress = 7 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = lMemory Or Memory(lMemoryAddress) * 2
                ElseIf lBitAddress = 9 Then
                    lBitAddress = 1
                Else
                    lMemory = lMemory \ 4
                End If
'                Debug.Print "R";
            Case 1 ' O
                lThisBitAddress = Adr And 7&
                lThisMemoryAddress = Adr \ 8&
                lBit = -((Memory(lThisMemoryAddress) And myLookup(lThisBitAddress)) > 0&)
                
                If mlOperandIndex = 0 Then
                    myFirstOperand = lBit
                    mlOperandIndex = 1
                Else
                    lBit = 1 - (myFirstOperand And lBit)
                    Memory(lThisMemoryAddress) = Memory(lThisMemoryAddress) And (myLookup(lThisBitAddress) Xor &HFF) Or lBit * myLookup(lThisBitAddress)
                    mlOperandIndex = 0
                End If
                
                lBitAddress = lBitAddress + 2
                If lBitAddress = 8 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = Memory(lMemoryAddress)
                    lBitAddress = 0
                ElseIf lBitAddress = 7 Then
                    lMemoryAddress = lMemoryAddress + 1
                    lMemory = lMemory Or Memory(lMemoryAddress) * 2
                ElseIf lBitAddress = 9 Then
                    lBitAddress = 1
                Else
                    lMemory = lMemory \ 4
                End If
'                Debug.Print "O";
            Case 3 ' J
'                Debug.Print "J";
                MsgBox Scientific(GetCounter)
                MsgBox (Memory(131072 \ 8))
                MsgBox (Memory(131080 \ 8))
                MsgBox Hex$(Memory(Temp() \ 8))

                lBitAddress = Adr And 7&
                lMemoryAddress = Adr \ 8&
                lMemory = Memory(lMemoryAddress) \ myLookup(lBitAddress)
                If lBitAddress = 7 Then
                    lMemory = lMemory Or Memory(lMemoryAddress + 1) * 2
                End If
                StartCounter
                
        End Select
    Loop Until False

End Sub
