Attribute VB_Name = "ArrayHandling"
Option Explicit

Public Function ArrayAppend(vArray As Variant, vItem As Variant)
    Dim lUbound As Long
    lUbound = UBound(vArray) + 1
    ReDim Preserve vArray(lUbound)
    vArray(lUbound) = vItem
End Function

Public Function Splice(vArray As Variant, ByVal lStart As Long, Optional lEnd As Long = -2) As Variant
    Dim vResult As Variant
    Dim lIndex As Long
    
    If lEnd = -2 Then
        lEnd = UBound(vArray)
    End If
    
    vResult = Array()
    
    If lEnd < lStart Then
        Splice = vResult
        Exit Function
    End If
    
    ReDim vResult(lEnd - lStart)
    
    For lIndex = lStart To lEnd
        vResult(lIndex - lStart) = vArray(lIndex)
    Next
    
    Splice = vResult
End Function

Public Function Concat(ParamArray vArrays() As Variant)
    Dim lArraysIndex As Long
    Dim lIndex As Long
    Dim vResult As Variant
    Dim lSize As Long
    Dim lCurrentIndex As Long
    
    For lArraysIndex = LBound(vArrays) To UBound(vArrays)
        lSize = lSize + UBound(vArrays(lArraysIndex)) - LBound(vArrays(lArraysIndex)) + 1
    Next
    lSize = lSize - 1
    
    vResult = Array()
    
    If lSize > -1 Then
        ReDim vResult(lSize)
        
        For lArraysIndex = LBound(vArrays) To UBound(vArrays)
            For lIndex = LBound(vArrays(lArraysIndex)) To UBound(vArrays(lArraysIndex))
                vResult(lCurrentIndex + lIndex - LBound(vArrays(lArraysIndex))) = vArrays(lArraysIndex)(lIndex)
            Next
            lCurrentIndex = lCurrentIndex + UBound(vArrays(lArraysIndex)) - LBound(vArrays(lArraysIndex)) + 1
        Next
    End If
    Concat = vResult
End Function

Public Sub Pad(vArray As Variant, lSize As Long, vPad As Variant)
    Dim lIndex As Long
    Dim lOriginalSize As Long
    
    lOriginalSize = UBound(vArray)
    
    ReDim Preserve vArray(LBound(vArray) To lSize)
    
    For lIndex = lOriginalSize + 1 To lSize
        vArray(lIndex) = vPad
    Next
End Sub

Public Function PaddedArray(lSize As Long, vPad As Variant) As Variant
    Dim lIndex As Long
    Dim vPaddedArray
    
    vPaddedArray = Array()
    
    If lSize > -1 Then
        ReDim Preserve vPaddedArray(lSize)
        
        For lIndex = 0 To lSize
            vPaddedArray(lIndex) = vPad
        Next
    End If
    PaddedArray = vPaddedArray
End Function

Public Sub Fill(vArray As Variant, vValue As Variant)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vArray)
        vArray(lIndex) = vValue
    Next
End Sub
