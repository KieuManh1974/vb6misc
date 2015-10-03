Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()
    Dim x As Long
    Dim k As String
    
    For x = 18 To 0 Step -1
        Debug.Print Key(x, 8) & " " & x
    Next
    k = Key(9, 8)
    End
    For x = 1 To 1000000
        If Val(Right$(Multiply(k, CStr(x) & "501"), 8)) = 1 Then
            Debug.Print x
        End If
    Next
End Sub

Public Function Multiply(ByVal A As String, ByVal B As String) As String
    Dim x As Long
    Dim y As Long
    Dim o As Long
    Dim p As Long
    Dim m As Long
    Dim s As Long
    Dim c As Long
    Dim r As Long
    
    If Len(B) > Len(A) Then
        A = String$(Len(B) - Len(A), "0") & A
    ElseIf Len(A) > Len(B) Then
        B = String$(Len(A) - Len(B), "0") & B
    End If
    r = Len(A)
    
    A = String$(Len(A), "0") & A
    B = String$(Len(B), "0") & B
    
    m = 2 * r - 2
    For o = 0 To m
        s = 0
        For x = 0 To o
            y = o - x
            s = s + Val(Mid$(A, Len(A) - x, 1)) * Val(Mid$(B, Len(B) - y, 1))
        Next
        'c = Left$(Format$(s, "00"), 1)
        Multiply = Right$(Format$(s, "00"), 1) & Multiply
    Next
End Function

Public Function Key(iIndex As Long, iLength As Long) As String
    Dim iX As Long
    
    For iX = 0 To iLength - 1
        Key = CStr((Factorial(iIndex + iX) / Factorial(iIndex) / Factorial(iX)) Mod 10) & Key
    Next
    
End Function

Private Function Factorial(iNumber As Long) As Double
    Dim iX As Long
    
    Factorial = 1
    For iX = 1 To iNumber
        Factorial = Factorial * iX
    Next
    
End Function

Public Function Reverse(sNumber As String) As String
    Dim iX As Long
    
    For iX = 1 To Len(sNumber)
        Reverse = Mid$(sNumber, iX, 1) & Reverse
    Next
End Function
