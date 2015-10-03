Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    FindSums
    'CountSeq
End Sub

Private Sub FindSums()
    Dim lList() As Long
    Dim bFound As Boolean
    Dim lNumberCheck As Long
    Dim lTest1 As Long
    Dim lTest2 As Long
    Dim lNewNumber As Long
    
    ReDim lList(0)
    lList(0) = 1
        
    For lNumberCheck = 2 To 50 Step 2
        bFound = False
        For lTest1 = 0 To UBound(lList)
            For lTest2 = 0 To UBound(lList)
                If (lList(lTest1) + lList(lTest2)) = lNumberCheck Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound Then
                Exit For
            End If
        Next
        If Not bFound Then
            lNewNumber = lNumberCheck - lList(UBound(lList))
            'lNewNumber = lNumberCheck - 1
            ReDim Preserve lList(UBound(lList) + 1)
            lList(UBound(lList)) = lNewNumber
        End If
    Next
End Sub

Private Sub CountSeq()
    Dim lSeq(32) As Long
    Dim lLoop1 As Long
    Dim lLoop2 As Long
    Dim vThisSet As Variant
    Dim lLoop3 As Long
    Dim bFound As Boolean
    
    Dim lLargestSet As Long
    Dim lMinimumMembers As Long
    Dim lMemberCount As Long
    
    Dim vSets As Variant
    Dim lCombinationIndex As Long
    
    Dim lBitIndex As Long
    
    Dim bSet(64) As Boolean
    Dim lIndex As Long
    Dim lContiguousCount As Long
    
    Dim lSetSize As Long
    
    vSets = Array()
    
    lMinimumMembers = 100
    
    For lCombinationIndex = 0 To 2 ^ 12 - 1
        lMemberCount = 0
        vThisSet = Array()
        For lLoop1 = 0 To 10
            If lSeq(lLoop1) = 1 Then
                lMemberCount = lMemberCount + 1
            End If
            For lLoop2 = 0 To 10
                If lSeq(lLoop1) = 1 And lSeq(lLoop2) = 1 Then
                    bFound = False
                    For lLoop3 = 0 To UBound(vThisSet)
                        If vThisSet(lLoop3) = (lLoop1 + lLoop2 + 2) Then
                            bFound = True
                            Exit For
                        End If
                    Next
                    If Not bFound Then
                        ReDim Preserve vThisSet(UBound(vThisSet) + 1)
                        vThisSet(UBound(vThisSet)) = (lLoop1 + lLoop2 + 2)
                    End If
                End If
            Next
        Next
        
        For lIndex = 0 To 64
            bSet(lIndex) = False
        Next
        For lIndex = 0 To UBound(vThisSet)
            bSet(vThisSet(lIndex)) = True
        Next

        lContiguousCount = 2
        While bSet(lContiguousCount)
            lContiguousCount = lContiguousCount + 2
        Wend
        
        'lSetSize = UBound(vThisSet) + 1
        lSetSize = lContiguousCount / 2 - 1
        
        If (lSetSize) > lLargestSet Then
            lLargestSet = lSetSize
            lMinimumMembers = lMemberCount
            vSets = Array(Array(lSeq, vThisSet))
        ElseIf (lSetSize) = lLargestSet Then
            If lMemberCount < lMinimumMembers Then
                vSets = Array(Array(lSeq, vThisSet))
                lMinimumMembers = lMemberCount
            ElseIf lMemberCount = lMinimumMembers Then
                ReDim Preserve vSets(UBound(vSets) + 1)
                vSets(UBound(vSets)) = Array(lSeq, vThisSet)
            End If
        End If
        
        lBitIndex = 0
        While lSeq(lBitIndex) = 1
            lSeq(lBitIndex) = 0
            lBitIndex = lBitIndex + 1
        Wend
        lSeq(lBitIndex) = 1
    Next
End Sub
