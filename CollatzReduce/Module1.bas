Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()
    Dim lNumber As Long
    
    For lNumber = 1 To 3399 Step 2
        Debug.Print lNumber & ":" & Reduce(lNumber)
    Next
End Sub

Public Function Reduce(ByVal lNumber As Long, ByVal lMultiplier) As Long
    Dim lAdd As Long
    Dim lSteps As Long
    Dim lOriginalNumber As Long
    Dim lMult As Long
    
    On Error GoTo exitreduce
    
    lOriginalNumber = DivideDown(lNumber)
    
    lMult = 1
    Do
        lNumber = lOriginalNumber
        lAdd = lMult
        Do
            lNumber = lNumber + lAdd
            lNumber = DivideDown(lNumber)
            If lNumber = 1 Then
                Reduce = lSteps
                Exit Function
            End If
            lAdd = lAdd / lMultiplier
        Loop Until lAdd = 0
        lMult = lMult * lMultiplier
        lSteps = lSteps + 1
    Loop
    
exitreduce:
    Reduce = -1000
End Function

Public Function DivideDown(ByVal lNumber As Long) As Long
    While (lNumber / 2) = Int(lNumber / 2)
        lNumber = lNumber / 2
    Wend
    DivideDown = lNumber
End Function
