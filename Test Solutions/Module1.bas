Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Sub main()
    Dim lDigits(3) As Long
    Dim lTempDigits(3) As Long
    Dim lOperators(2) As Long
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lIndex3 As Long
    
    lDigits(0) = 1
    lDigits(1) = 2
    lDigits(2) = 3
    lDigits(3) = 5
    
    For lIndex2 = 0 To 4 ^ 3 - 1
        CountDigits lOperators, lIndex2, 4
        For lIndex = 0 To 23
            CopyMemory lTempDigits(0), lDigits(0), 16&
            PermuteDigits lTempDigits, lIndex
        Next
        Debug.Print Value(lTempDigits, lOperators)
'        If Value(lTempDigits, lOperators) = 100 Then
'            Debug.Print Expression(lTempDigits, lOperators)
'        End If
    Next
End Sub


Public Sub PermuteDigits(lDigits() As Long, ByVal lPermutation As Long)
    Dim lRotations() As Long
    Dim lMod As Long
    Dim lIndex As Long
    Dim lTemp() As Long
    Dim lDigit As Long
    Dim lSize As Long
    
    lSize = UBound(lDigits)
    
    ReDim lRotations(lSize)
    ReDim lTemp(lSize)
    
    lMod = 2
    While lIndex <= lSize
        lRotations(lIndex) = lPermutation Mod lMod
        lPermutation = lPermutation \ lMod
        lMod = lMod + 1
        lIndex = lIndex + 1
    Wend
    
    For lIndex = lSize To 1 Step -1
        For lDigit = lIndex To 0 Step -1
            lTemp(lDigit) = lDigits((lDigit - lRotations(lIndex - 1) + lIndex + 1) Mod (lIndex + 1))
        Next
        For lDigit = lIndex To 0 Step -1
            lDigits(lDigit) = lTemp(lDigit)
        Next
    Next
End Sub


Public Sub CountDigits(lDigits() As Long, ByVal lCombination As Long, ByVal lBase As Long)
    Dim lIndex As Long
    Dim lMod As Long
    
    lMod = lBase
    
    For lIndex = 0 To UBound(lDigits)
        lDigits(lIndex) = lCombination Mod lMod
        lCombination = lCombination \ lBase
    Next
End Sub

Public Function Value(lDigits() As Long, lOperators() As Long) As Double
    Dim fRunning As Double
    Dim lIndex As Long
    
    fRunning = lDigits(0)
    For lIndex = 1 To UBound(lDigits)
        Select Case lOperators(lIndex - 1)
            Case 0
                fRunning = fRunning + lDigits(lIndex)
            Case 1
                fRunning = fRunning - lDigits(lIndex)
            Case 2
                fRunning = fRunning * lDigits(lIndex)
            Case 3
                fRunning = fRunning / lDigits(lIndex)
        End Select
    Next
    Value = fRunning
End Function

Public Function Expression(lDigits() As Long, lOperators() As Long) As String
    Dim fRunning As Double
    Dim lIndex As Long
    
    Expression = lDigits(0)
    For lIndex = 1 To UBound(lDigits)
        Select Case lOperators(lIndex - 1)
            Case 0
                Expression = Expression & "+" & lDigits(lIndex)
            Case 1
                Expression = Expression & "-" & lDigits(lIndex)
            Case 2
                Expression = Expression & "*" & lDigits(lIndex)
            Case 3
                Expression = Expression & "/" & lDigits(lIndex)
        End Select
    Next
End Function
