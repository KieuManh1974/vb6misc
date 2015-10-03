Attribute VB_Name = "Module1"
Option Explicit

Public Sub ThreeOnes(ByVal sPattern As String)
    Dim lSize As Long
    Dim lCheck As Long
    Dim sModuliToSort() As String
    Dim lModuliToSortCount As Long
    Dim lModCheck As Long
    
    lSize = Len(sPattern) \ 2 ' integer division i.e. floor of half length
    
    ' Step 1: Scan for ones and find the moduli and residues
    For lCheck = 1 To Len(sPattern)
        If Mid$(sPattern, lCheck, 1) = "1" Then
            For lModCheck = 1 To lSize ' Only actually need to check primes
                ReDim Preserve sModuliToSort(lModuliToSortCount)
                sModuliToSort(sModuliToSort) = lModCheck & ":" & (lCheck - 1) Mod lModCheck
                lModuliToSortCount = lModuliToSortCount + 1
            Next
        End If
    Next
    
    ' Step 2: Sort sModuliToSort array in n * log (n) time
    ' Step 3: Scan through and find three consecutive matches
End Sub

Sub main()
    Dim lSize As Long
    
    For lSize = 1 To 20
        Debug.Print CheckString(lSize) + 1
    Next
End Sub

Public Function CheckString(ByVal lNum As Long) As Long
    Dim lArray() As Long
    Dim lColumn As Long
    Dim bFinished As Boolean
    
    ReDim lArray(lNum - 1)
    
    Do
        lColumn = -1
        Do
            lColumn = lColumn + 1
            If lColumn = lNum Then
                bFinished = True
                Exit Do
            End If
            lArray(lColumn) = 1 - lArray(lColumn)
        Loop Until lArray(lColumn) = 1
        
        If bFinished Then Exit Function
        
        If Not SpacedFound(lArray) Then
            CheckString = CheckString + 1
'            For lColumn = 0 To lNum - 1
'                Debug.Print lArray(lColumn);
'            Next
            'Debug.Print
        End If
    Loop
End Function

Public Function SpacedFound(lArray() As Long) As Boolean
    Dim lOnes() As Long
    Dim lOneCount As Long
    Dim lColumn As Long
    Dim lCheck1 As Long
    Dim lCheck2 As Long
    Dim lSpacing As Long
    
    For lColumn = 0 To UBound(lArray)
        If lArray(lColumn) = 1 Then
            ReDim Preserve lOnes(lOneCount)
            lOnes(lOneCount) = lColumn
            lOneCount = lOneCount + 1
        End If
    Next
    
    If lOneCount < 3 Then
        SpacedFound = False
        Exit Function
    End If
    
    For lCheck1 = 0 To lOneCount - 1
        For lCheck2 = lCheck1 + 1 To lOneCount - 1
            lSpacing = lOnes(lCheck2) - lOnes(lCheck1)
            If (lOnes(lCheck2) + lSpacing) <= UBound(lArray) Then
                If lArray(lOnes(lCheck2) + lSpacing) = 1 Then
                    SpacedFound = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function
