Attribute VB_Name = "Compiler"
Option Explicit

Public Sub Add(oSymbol1 As clsSymbol, oSymbol2 As clsSymbol)
    Dim oResult As clsSymbol
    Dim lByte As Long
    
    Set oResult = CreateSymbol(AddSize(oSymbol1.ByteSize, oSymbol2.ByteSize))
    oSymbolTable.AddSymbol oResult
    
    oOutput.AddOutput "CLC"
    For lByte = 0 To oResult.ByteSize - 2
        oOutput.AddOutput "LDA " & oSymbol1.AddressLocation + lByte
        oOutput.AddOutput "ADC " & oSymbol2.AddressLocation + lByte
        oOutput.AddOutput "STA " & oResult.AddressLocation + lByte
    Next
End Sub

Public Sub Copy(oSymbol1 As clsSymbol, oSymbol2 As clsSymbol)
    Dim lByte As Long
    
    For lByte = 0 To oSymbol1.ByteSize - 1
        oOutput.AddOutput "LDA " & oSymbol1.AddressLocation + lByte
        oOutput.AddOutput "STA " & oSymbol2.AddressLocation + lByte
    Next
End Sub


Public Function AddSize(lSize1 As Long, lSize2 As Long)
    If lSize1 >= lSize2 Then
        AddSize = lSize1 + 1
    Else
        AddSize = lSize2 + 1
    End If
End Function


Public Function CreateSymbol(lByteSize As Long) As clsSymbol
    Dim oRange As New clsRange
    Dim oField As New clsField
    Dim oUnit As New clsUnit
    Dim oSymbol As New clsSymbol
    
    oRange.PhysicalStart = 0
    oRange.PhysicalEnd = 2 ^ (lByteSize * 8) - 1
    
    oField.AddRange oRange
    
    oUnit.AddField oField
    
    Set oSymbol.Unit = oUnit
    oSymbol.SymbolType = stVar
    
    Set CreateSymbol = oSymbol
End Function
