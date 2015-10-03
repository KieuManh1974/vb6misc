Attribute VB_Name = "Module1"
Sub main()
    BinarySearch
End Sub

Sub BinarySearch()
    Dim lElements() As Long
    Dim lElementCount As Long
    Dim lNumber As Long
    Dim lPosition As Long
    Dim bFound As Boolean
    Dim lIndex As Long
    Dim bFinished As Boolean
    Dim lHigh As Long
    Dim lLow As Long
    Dim lMid As Long
    Dim lDirection As Long
    
    Randomize 50
    For x = 0 To 10000
        lNumber = CInt(Rnd * 500)
        
        lHigh = lElementCount - 1
        lLow = 0
        
        bFound = False
        bFinished = False
        If lElementCount <> 0 Then
            lMid = Int((lLow + lHigh) / 2)
        
            While Not bFound And Not bFinished
                If lElements(lMid) = lNumber Then
                    bFound = True
                    lPosition = lMid
                ElseIf lLow = lHigh Then
                    bFinished = True
                    lPosition = lLow
                    If lNumber < lElements(lLow) Then
                        lDirection = 0
                    Else
                        lDirection = 1
                    End If
                ElseIf lLow = (lHigh - 1) Then
                    bFinished = True
                    If lNumber < lElements(lLow) Then
                        lDirection = 0
                        lPosition = lLow
                    ElseIf lNumber > lElements(lHigh) Then
                        lPosition = lHigh
                        lDirection = 1
                    ElseIf lNumber = lElements(lHigh) Then
                        bFound = True
                        lPosition = lHigh
                    ElseIf lNumber > lElements(lLow) Then
                        lPosition = lLow
                        lDirection = 1
                    End If
                ElseIf lNumber > lElements(lMid) Then
                    lLow = lMid
                    lMid = Int((lLow + lHigh) / 2)
                Else
                    lHigh = lMid
                    lMid = Int((lLow + lHigh) / 2)
                End If
            Wend
        End If
        
        If Not bFound Then
            ReDim Preserve lElements(lElementCount)
            
            For lIndex = lElementCount To lPosition + 1 + lDirection Step -1
                lElements(lIndex) = lElements(lIndex - 1)
            Next
            lElements(lPosition + lDirection) = lNumber
            lElementCount = lElementCount + 1
        End If
    Next
    
End Sub
