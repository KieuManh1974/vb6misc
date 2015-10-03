Attribute VB_Name = "Program"
Option Explicit

Public oSymbolTable As New clsSymbolTable
Public oOutput As New clsOutput

Sub main()
    Compile
End Sub

Sub Compile()
    Dim oRange As New clsRange
    Dim oField As New clsField
    Dim oUnit As New clsUnit
    Dim oSymbol1 As New clsSymbol
    Dim oSymbol2 As New clsSymbol

    
    oRange.RangeType = rtRange
    oRange.PhysicalStart = 0
    oRange.PhysicalEnd = 255
    
    oField.AddRange oRange
    oField.MaxValue = 255
    
    oUnit.AddField oField
    
    oSymbol1.Name = "a"
    oSymbol1.SymbolType = stVar
    Set oSymbol1.Unit = oUnit
    oSymbolTable.AddSymbol oSymbol1
    
    oSymbol2.Name = "b"
    oSymbol2.SymbolType = stVar
    Set oSymbol2.Unit = oUnit
    oSymbolTable.AddSymbol oSymbol2
    
    Compiler.Add oSymbol1, oSymbol2
End Sub
