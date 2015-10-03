Attribute VB_Name = "Module1"
Option Explicit

Sub main()

End Sub

Public Function CollatzSize(ByVal x As Double) As Double
    Dim dMax As Double
    Dim dVal As Double
    Dim dSteps As Double
    Dim dStep As Double
    
    While x > 1
        dVal = Log(x) / Log(2)
        If dVal > dMax Then
            dMax = dVal
            dSteps = dStep
        End If
        x = Reduce((3 * x) + 1)
        dStep = dStep + 1
    Wend
    CollatzSize = dSteps
End Function

Public Function CollatzSteps(ByVal x As Long) As Long
    Dim dMax As Double
    Dim dVal As Double
    Dim dSteps As Double
    Dim dStep As Double
    
    While x > 1
        If x Mod 2 = 0 Then
            x = OneStep(x)
        Else
            x = x * 3& + 1&
        End If
        dStep = dStep + 1
    Wend
    CollatzSteps = dStep
End Function

Public Function OneStep(ByVal x As Long) As Long
    While x Mod 2 = 0
        x = x \ 2
    Wend
    OneStep = x
End Function

Public Function Add1(ByVal y As Long) As Long
    Dim p As Long
    Dim q As Long
    
    p = 1
    While (y Mod 2) = 0
        p = p * 2
        y = y \ 2
    Wend
    Add1 = (y + 1) * p
    
End Function

Public Function Reduce(ByVal y As Double) As Double
    Dim p As Long
    Dim q As Long
    
    p = 1
    While (y / 2) - Int(y / 2) = 0
        p = p * 2
        y = y / 2
    Wend
    Reduce = y
End Function

Public Function Binary(ByVal x As Long) As String
    Dim lSize As Long
    Dim lIndex As Long
    
    lSize = Int(Log(x) / Log(2))
    
    For lIndex = 0 To lSize
        Binary = (x Mod 2) & Binary
        x = x \ 2
    Next
End Function

Public Function ZeroCount(ByVal x As Long) As Long
    Dim sBinary As String
    Dim lIndex As Long
    
    sBinary = Binary(x)
    
    For lIndex = 1 To Len(sBinary)
        If Mid$(sBinary, lIndex, 1) = "0" Then
            ZeroCount = ZeroCount + 1
        End If
    Next
End Function
