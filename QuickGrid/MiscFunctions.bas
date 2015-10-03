Attribute VB_Name = "MiscFunctions"
Option Explicit

Public Function Largest(ParamArray vValues() As Variant)
    Dim lIndex As Long
    
    For lIndex = LBound(vValues) To UBound(vValues)
        If vValues(lIndex) > Largest Then
            Largest = vValues(lIndex)
        End If
    Next
End Function

