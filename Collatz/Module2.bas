Attribute VB_Name = "Module2"
Option Explicit

Public Sub Test()
    Dim vMachine1 As Variant
    Dim vCombined As Variant
    
    vMachine1 = Array("A C A B", "B N C B", "C I D B", "D N D C")
    vCombined = MachineCombiner(vMachine1, vMachine1)
End Sub

Public Function MachineCombiner(vMachine1 As Variant, vMachine2 As Variant) As Variant
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim sName(2) As String
    Dim sType(2) As String
    Dim sNextZeroState(2) As String
    Dim sNextOneState(2) As String
    Dim vSplit(2) As Variant
    Dim vMachineOut As Variant
    
    vMachineOut = Array()
    
    For lIndex1 = 0 To UBound(vMachine1)
        vSplit(0) = Split(vMachine1(lIndex1), " ")
        sName(0) = vSplit(0)(0)
        sType(0) = vSplit(0)(1)
        sNextZeroState(0) = vSplit(0)(2)
        sNextOneState(0) = vSplit(0)(3)
        
        For lIndex2 = 0 To UBound(vMachine2)
            vSplit(1) = Split(vMachine1(lIndex2), " ")
            sName(1) = vSplit(1)(0)
            sType(1) = vSplit(1)(1)
            sNextZeroState(1) = vSplit(1)(2)
            sNextOneState(1) = vSplit(1)(3)
            
            sName(2) = sName(0) & sName(1)
            If sType(1) = "C" Then
              sType(2) = "C"
            ElseIf sType(1) = "S" Then
                sType(2) = "S"
            ElseIf sType(1) = "N" Then
                sType(2) = sType(0)
            ElseIf sType(1) = "I" Then
                Select Case sType(0)
                    Case "C"
                        sType(2) = "S"
                    Case "S"
                        sType(2) = "C"
                    Case "N"
                        sType(2) = "I"
                    Case "I"
                        sType(2) = "N"
                End Select
            End If
            
            sNextZeroState(2) = sNextZeroState(0) & sNextZeroState(1)
            sNextOneState(2) = sNextOneState(0) & sNextOneState(1)
            
            ReDim Preserve vMachineOut(UBound(vMachineOut) + 1)
            vMachineOut(UBound(vMachineOut)) = sName(2) & " " & sType(2) & " " & sNextZeroState(2) & " " & sNextOneState(2)
        Next
    Next
    
    MachineCombiner = vMachineOut
End Function
