Attribute VB_Name = "Arithmetic"
Option Explicit

Public Const SYMBOLS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Function Convert(sNumber As String, iBase1 As Integer, iBase2 As Integer) As String
    Dim iDigitIndex As Integer
    Dim iDigit As Integer
    Dim sMultiplier As String
    
    Convert = "0"
    sMultiplier = ConvertDigit(iBase1, iBase2)
    For iDigitIndex = 1 To Len(sNumber)
        iDigit = InStr(SYMBOLS, Mid$(sNumber, iDigitIndex, 1)) - 1
        
        Convert = Multiply(Convert, sMultiplier, iBase2)
        Convert = Add(Convert, ConvertDigit(iDigit, iBase2), iBase2)
    Next
End Function

Private Function ConvertDigit(iDigit As Integer, iBase As Integer) As String
    Dim sUnit As String
    Dim sTens As String
    
    If iDigit = 0 Then
        ConvertDigit = "0"
    Else
        While iDigit > 0
            ConvertDigit = Mid$(SYMBOLS, (iDigit Mod iBase) + 1, 1) & ConvertDigit
            iDigit = iDigit \ iBase
        Wend
    End If
End Function

Public Function Multiply(sNumber1 As String, sNumber2 As String, iBase As Integer, Optional iNoCarry As Integer = 1) As String
    Dim iIndex1 As Integer
    Dim iIndex2 As Integer
    Dim iDigit1 As Integer
    Dim iDigit2 As Integer
    Dim iCarry As Integer
    Dim iSum As Integer
    Dim iMult As Integer
    Dim sMult As String
    
    Multiply = "0"
    For iIndex1 = Len(sNumber1) To 1 Step -1
        iDigit1 = InStr(SYMBOLS, Mid$(sNumber1, iIndex1, 1)) - 1
        iCarry = 0
        sMult = ""
        For iIndex2 = Len(sNumber2) To 1 Step -1
            iDigit2 = InStr(SYMBOLS, Mid$(sNumber2, iIndex2, 1)) - 1
            iMult = iDigit1 * iDigit2 + iCarry * iNoCarry
            
           iSum = iMult Mod iBase
           iCarry = iMult \ iBase
           
           sMult = Mid$(SYMBOLS, iSum + 1, 1) & sMult
        Next
        sMult = Mid$(SYMBOLS, iCarry * iNoCarry + 1, 1) & sMult
        sMult = sMult & String$(Len(sNumber1) - iIndex1, "0")
        Multiply = Add(Multiply, sMult, iBase, iNoCarry)
    Next
End Function

Public Function Add(sNumber1 As String, sNumber2 As String, iBase As Integer, Optional iNoCarry As Integer = 1) As String
    Dim iIndex As Integer
    Dim iCarry As Integer
    Dim iDigit1 As Integer
    Dim iDigit2 As Integer
    Dim iSum As Integer
    
    If Len(sNumber1) > Len(sNumber2) Then
        sNumber2 = String$(Len(sNumber1) - Len(sNumber2), "0") & sNumber2
    ElseIf Len(sNumber1) < Len(sNumber2) Then
        sNumber1 = String$(Len(sNumber2) - Len(sNumber1), "0") & sNumber1
    End If
    
    For iIndex = Len(sNumber1) To 1 Step -1
        iDigit1 = InStr(SYMBOLS, Mid$(sNumber1, iIndex, 1)) - 1
        iDigit2 = InStr(SYMBOLS, Mid$(sNumber2, iIndex, 1)) - 1
        
        iSum = (iDigit1 + iDigit2 + iCarry * iNoCarry) Mod iBase
        iCarry = (iDigit1 + iDigit2 + iCarry * iNoCarry) \ iBase
        Add = Mid$(SYMBOLS, iSum + 1, 1) & Add
    Next
    Add = Mid$(SYMBOLS, iCarry * iNoCarry + 1, 1) & Add
    Add = StripZero(Add)
End Function

Public Function Subs(sNumber1 As String, sNumber2 As String, iBase As Integer, Optional iNoCarry As Integer = 1) As String
    Dim iIndex As Integer
    Dim iCarry As Integer
    Dim iDigit1 As Integer
    Dim iDigit2 As Integer
    Dim iSum As Integer
    
    If Len(sNumber1) > Len(sNumber2) Then
        sNumber2 = String$(Len(sNumber1) - Len(sNumber2), "0") & sNumber2
    ElseIf Len(sNumber1) < Len(sNumber2) Then
        sNumber1 = String$(Len(sNumber2) - Len(sNumber1), "0") & sNumber1
    End If
    
    For iIndex = Len(sNumber1) To 1 Step -1
        iDigit1 = InStr(SYMBOLS, Mid$(sNumber1, iIndex, 1)) - 1
        iDigit2 = InStr(SYMBOLS, Mid$(sNumber2, iIndex, 1)) - 1
        
        iSum = (iDigit1 - iDigit2 - iCarry * iNoCarry + iBase) Mod iBase
        If (iDigit1 - iDigit2 - iCarry * iNoCarry) < 0 Then
            iCarry = 1
        Else
            iCarry = 0
        End If

        Subs = Mid$(SYMBOLS, iSum + 1, 1) & Subs
    Next
    Subs = Mid$(SYMBOLS, iCarry * iNoCarry + 1, 1) & Subs
    Subs = StripZero(Subs)
End Function

Public Function StripZero(sNumber As String) As String
    Dim bNotIgnore As Boolean
    Dim sDigit As String
    Dim iIndex As Integer
    
    For iIndex = 1 To Len(sNumber)
        sDigit = Mid$(sNumber, iIndex, 1)
        If Not bNotIgnore Then
            If sDigit <> "0" Then
                bNotIgnore = True
                StripZero = StripZero & sDigit
            End If
        Else
            StripZero = StripZero & sDigit
        End If
    Next
    If StripZero = "" Then
        StripZero = "0"
    End If
End Function

Public Function Reverse(number As String) As String
    Dim x As Long
    
    For x = 1 To Len(number)
        Reverse = Mid(number, x, 1) & Reverse
    Next
End Function

Public Function Pad(sNumber As String, iSize) As String
    If iSize > Len(sNumber) Then
        Pad = String$(iSize - Len(sNumber), "0") & sNumber
    Else
        Pad = Right$(sNumber, iSize)
    End If
End Function

Public Function Slice(sString As String, iStart As Integer, Optional iStop As Integer) As String
    If iStop = 0 Then
        Slice = Mid$(sString, iStart, 1)
    Else
        Slice = Mid$(sString, iStart, iStop - iStart + 1)
    End If
End Function

Public Function Factorial(A As Long) As Double
    Dim i As Long
    
    Factorial = 1
    For i = 1 To A
        Factorial = Factorial * i
    Next
End Function

Public Function GenerateKey(ByVal iValue As Long, ByVal iLength As Long) As String
    Dim iIndex As Long
    
    For iIndex = 0 To iLength - 1
        GenerateKey = CStr(Factorial(iValue + iIndex) / Factorial(iValue) / Factorial(iIndex) Mod 10) & GenerateKey
    Next
End Function


