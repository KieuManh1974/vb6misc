Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim vOrdered As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lNumber As Long
    Dim vArray As Variant
    Dim lSize As Long
    Dim sOrder As String
    Dim lSpacing As Long
    Dim lPower As Long
    Dim lLevel As Long
    Dim lPosition As Long
    Dim lUnique(1000000#) As Long
    Dim sPermutation(10000#) As String
    Dim sNumber As String
    
    vArray = Array("0", "1", "2", "3")
    lSize = UBound(vArray) + 1
        
    For lIndex = 0 To Factorial(lSize) - 1
        vOrdered = Permute(vArray, lIndex)
        lNumber = 0
        lPower = 0
        sNumber = ""
        
        For lLevel = 0 To lSize - 2
            For lSpacing = 1 To lLevel + 1
                lPosition = lLevel - lSpacing + 1
                If vOrdered(lPosition) > vOrdered(lPosition + lSpacing) Then
                    lNumber = lNumber + 2 ^ lPower
                    sNumber = sNumber & "1"
                Else
                    sNumber = sNumber & "0"
                End If
                lPower = lPower + 1
                If lPower = 5 Then Exit For
            Next
            If lPower = 5 Then Exit For
        Next
        
        lUnique(lNumber) = lUnique(lNumber) + 1
        sOrder = ""
        For lIndex2 = 0 To lSize - 1 Step 1
            sOrder = sOrder & vOrdered(lIndex2)
        Next
        sPermutation(lNumber) = sPermutation(lNumber) & " " & sOrder
        Debug.Print lIndex & " " & sOrder & " " & lNumber & " "; sNumber
    Next
    Debug.Print
    Dim lTotal As Long
    lTotal = 0
    For lIndex2 = 0 To 2 ^ lPower - 1
        If lUnique(lIndex2) = 1 Then Debug.Print lIndex2 & " " & sPermutation(lIndex2): lTotal = lTotal + 1
    Next
    'Debug.Print lTotal
End Sub

Private Function Permute(ByVal vList As Variant, ByVal lPermutationNumber As Long) As Variant
    Dim lSize As Long
    Dim lPositions As Long
    Dim lDivider As Long
    Dim lPosition As Long
    Dim vPositions As Variant
    Dim lUpdatePosition As Long
    Dim lIndex As Long
    Dim vOut As Variant
    Dim lModulus As Long
    
    vPositions = Array()
    vOut = Array()
    
    lSize = UBound(vList) - LBound(vList) + 1
    
    ReDim vPositions(lSize - 1)
    ReDim vOut(lSize - 1)
    
    For lPosition = 0 To lSize - 1
        vPositions(lPosition) = lPosition
    Next
    For lPosition = lSize - 2 To 0 Step -1
        lDivider = Factorial(lPosition + 1)
        lModulus = lPosition + 2
        lIndex = (lPermutationNumber \ lDivider) Mod lModulus
        For lUpdatePosition = 0 To lSize - 1
            If vPositions(lUpdatePosition) <= (lPosition + 1) Then
                vPositions(lUpdatePosition) = (vPositions(lUpdatePosition) + lIndex) Mod (lModulus)
            End If
        Next
    Next
    
    For lPosition = 0 To lSize - 1
        vOut(vPositions(lPosition)) = vList(lPosition)
    Next
    
    Permute = vOut
End Function

Public Function Factorial(ByVal lValue As Long) As Long
    Dim lIndex As Long
    
    Factorial = 1
    For lIndex = 2 To lValue
        Factorial = Factorial * lIndex
    Next
End Function
