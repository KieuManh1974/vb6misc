Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim phiA As Double
    Dim phiB As Double
    Dim lX As Long
    Dim fS As Double
    
    phiA = (Sqr(5) + 1) / 2
    phiB = (Sqr(5) - 1) / 2
    
    For lX = 0 To 5
        fS = fS + phiA ^ lX - (-phiB) ^ lX
'        fS = fS + phiA ^ lX
'        fS = fS + (-phiB) ^ lX
    Next
    Debug.Print fS
    
   ' Debug.Print ((phiA ^ 6) / (phiA - 1) - ((-phiB) ^ 6) / (-phiB - 1)) - ((phiA ^ 0 / (phiA - 1) - (-phiB) ^ 0) / ((-phiB) - 1))
    Debug.Print (((phiA ^ 6) / (phiA - 1)) - (phiA ^ 0 / (phiA - 1))) - (((-phiB) ^ 6 / (-phiB - 1)) - ((-phiB) ^ 0 / ((-phiB) - 1)))
    'Debug.Print ((-phiB) ^ 6 / (-phiB - 1)) - ((-phiB) ^ 0 / ((-phiB) - 1))
    
End Sub

Public Function Fib(ByVal x As Long) As Double
    Dim lIndex As Long
    Dim lSum As Long
    
    For lIndex = 1 To x Step 2
        lSum = lSum + Fac(x) / Fac(lIndex) / Fac(x - lIndex) * 5 ^ ((lIndex - 1) / 2)
    Next
    Fib = 2 * lSum / 2 ^ x
End Function

Public Function Fac(ByVal n As Long) As Long
    Dim lProduct As Long
    Dim lIndex As Long
    
    lProduct = 1
    
    For lIndex = 2 To n
        lProduct = lProduct * lIndex
    Next
    Fac = lProduct
End Function

