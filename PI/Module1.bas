Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim total As Currency
    Dim accumulator As Long
    Dim x As Long
    Dim y As Long
    Dim north As Long
    Dim west As Long
    
    x = 100000
    
    While x > 0
        north = accumulator + 2 * y + 1
        west = accumulator - 2 * x + 1
        
        If Abs(north) <= Abs(west) Then
            accumulator = north
            y = y + 1
            total = total + x
        Else
            accumulator = west
            x = x - 1
        End If
    Wend
    Debug.Print total
    Debug.Print Atn(1)
End Sub
