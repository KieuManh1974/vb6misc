Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim x As Long
    Dim r As Long
    Dim t As Long
    
    t = 2 ^ 8
    For x = 1 To t
        If IsPower(Polynomial(x)) Then
            r = r + 1
            'Debug.Print x
        End If
    Next
    Debug.Print r / t
End Sub

Private Function Add(ByVal lNumber As Long, ByVal lAddendum As Long) As Long
    Dim lDivide As Long
    
    lDivide = 1
    While Not (lNumber Mod 2) = 1
        lNumber = lNumber \ 2
        lDivide = lDivide * 2
    Wend
    
    Add = (lNumber + lAddendum) * lDivide
End Function

Private Function IsPower(ByVal lNumber As Long)
    While Not (lNumber Mod 2) = 1
        lNumber = lNumber \ 2
    Wend
    IsPower = lNumber = 1
End Function

Private Function Polynomial(ByVal lNumber As Long) As Long
    Dim lPower As Long
    
    lPower = 3 ^ 9
    
    While lPower > 0
        lNumber = Add(lNumber, lPower)
        lPower = lPower \ 3
    Wend
    Polynomial = lNumber
End Function
