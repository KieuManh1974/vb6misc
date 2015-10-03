Attribute VB_Name = "CompareLists"
Option Explicit

Private aList1() As Variant
Private aList2() As Variant

Public Sub CompareList()
    Dim lListIndex1 As Long
    Dim lListIndex2 As Long
    Dim lPosition1 As Long
    Dim lPosition2 As Long
    Dim lIndex As Long
    Dim bMatchCondition1 As Boolean
    Dim bMatchCondition2 As Boolean
    
    Initialise
    
    Do
        If lListIndex1 > UBound(aList1) Then
            Debug.Print "2: " & aList2(lListIndex2)
            lListIndex2 = lListIndex2 + 1
        ElseIf lListIndex2 > UBound(aList2) Then
            Debug.Print "1: " & aList1(lListIndex1)
            lListIndex1 = lListIndex1 + 1
        ElseIf aList1(lListIndex1) = aList2(lListIndex2) Then
            lListIndex1 = lListIndex1 + 1
            lListIndex2 = lListIndex2 + 1
        ElseIf aList1(lListIndex1) > aList2(lListIndex2) Then
            Debug.Print "2: " & aList2(lListIndex2)
            lListIndex2 = lListIndex2 + 1
        ElseIf aList1(lListIndex1) < aList2(lListIndex2) Then
            Debug.Print "1: " & aList1(lListIndex1)
            lListIndex1 = lListIndex1 + 1
        End If
    Loop Until lListIndex1 > UBound(aList1) And lListIndex2 > UBound(aList2)

End Sub

Private Sub Initialise()
    aList1 = Array("cat", "dog", "rabbit", "rabbit", "zoo")
    aList2 = Array("cat", "dog", "horse", "rabbit", "zog")
End Sub
