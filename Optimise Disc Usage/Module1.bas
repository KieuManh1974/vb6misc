Attribute VB_Name = "Module1"
Option Explicit
 
Public Sub main()
    Allocate 85, 5, Array(2, 85, 13, 50)
End Sub
 
Public Function Allocate(lSize As Long, lCount As Long, vSizes As Variant) As Variant
    Dim lAllocation() As Long
    Dim lReturnAllocation() As Long
    Dim lGap As Long
    Dim lMinUnits As Long
    Dim lUnitRemaining() As Long
    Dim bOk As Boolean
    Dim lUnitIndex As Long
    Dim lFileIndex As Long
    Dim lMaxCheckIndex As Long
    Dim lMinGap As Long
    
    Dim lFileCount As Long
    Dim lThisSize As Long
    Dim lMinCombinationCount As Long
    Dim lCombinationIndex As Long
    Dim bAdd As Boolean
    Dim lCombinationCount As Long
    
    lFileCount = UBound(vSizes)
    
    ReDim lAllocation(lFileCount) As Long
    ReDim lReturnAllocation(lFileCount) As Long
    ReDim lUnitRemaining(lCount - 1) As Long
 
    lMinUnits = lCount
    lMinGap = lSize * lCount
    
    For lFileIndex = 0 To lFileCount
        lReturnAllocation(lFileIndex) = -1
    Next
    
    For lCombinationIndex = 0 To lCount ^ (lFileCount + 1) - 1
        For lUnitIndex = 0 To lCount - 1
            lUnitRemaining(lUnitIndex) = lSize
        Next
    
        For lFileIndex = 0 To lFileCount
            lUnitRemaining(lAllocation(lFileIndex)) = lUnitRemaining(lAllocation(lFileIndex)) - vSizes(lFileIndex)
        Next
        
        bOk = True
        For lUnitIndex = 0 To lCount - 1
            If lUnitRemaining(lUnitIndex) < 0 Then
                bOk = False
                Exit For
            End If
        Next
        
        If bOk Then
            lMaxCheckIndex = lCount - 1
            Do While lMaxCheckIndex >= 0
                If lUnitRemaining(lMaxCheckIndex) <> lSize Then
                    Exit Do
                End If
                lMaxCheckIndex = lMaxCheckIndex - 1
            Loop
 
            If lMaxCheckIndex <= lMinUnits Then
                lGap = 0
                For lUnitIndex = 0 To lMaxCheckIndex - 1
                    lGap = lGap + lUnitRemaining(lUnitIndex)
                Next
            
                lMinUnits = lMaxCheckIndex
                If lGap < lMinGap Then
                    lMinGap = lGap
                    lReturnAllocation = lAllocation
                End If
            End If
        End If
        
        For lFileIndex = 0 To lFileCount
            lAllocation(lFileIndex) = (lAllocation(lFileIndex) + 1) Mod lCount
            If lAllocation(lFileIndex) <> 0 Then
                Exit For
            End If
        Next
    Next
    
    
End Function

