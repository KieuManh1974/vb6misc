Attribute VB_Name = "Module1"
Option Explicit

Public Const SYMBOLS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Sub main()
    TestCarryCombinations
End Sub

Private Sub TestCarryCombinations()
    Dim lNumber1 As Long
    Dim lNumber2 As Long
    Dim lCounter As Long
    Dim lCarryCount As Long
    Dim lBase1 As Long
    Dim lBase2 As Long
    Dim lMinCount As Long
    Dim lLimit As Long
    
    lMinCount = 2 ^ 30
    lLimit = 31
    
    For lBase1 = 31 To 2 Step -1
        For lBase2 = 31 To lBase1 Step -1
            lCarryCount = 0
            lCounter = 0
            For lNumber1 = 0 To lLimit
                For lNumber2 = lNumber1 To lLimit
                    If CheckForCarryMain(lNumber1, lNumber2, lBase1) And CheckForCarryMain(lNumber1, lNumber2, lBase2) Then
                        lCarryCount = lCarryCount + 1
                    End If
                    
                    lCounter = lCounter + 1
                Next
            Next
            If lCarryCount < lMinCount Then
                lMinCount = lCarryCount
                Debug.Print lBase1 & " " & lBase2 & " " & lMinCount & "/" & lCounter & " : " & lMinCount / lCounter
            End If
        Next
    Next
End Sub

Private Function CheckForCarryMain(ByVal lNumber1 As Long, ByVal lNumber2 As Long, ByVal lBase As Long) As Boolean
    Dim sNumber1 As String
    Dim sNumber2 As String
    
    sNumber1 = Convert(CStr(lNumber1), 10, lBase)
    sNumber2 = Convert(CStr(lNumber2), 10, lBase)
    
    If Len(sNumber1) > Len(sNumber2) Then
        sNumber2 = Pad(sNumber2, Len(sNumber1))
    Else
        sNumber1 = Pad(sNumber1, Len(sNumber2))
    End If
    
    CheckForCarryMain = CheckForCarry(sNumber1, sNumber2, lBase)
End Function

Private Function CheckForCarry(sNumber1 As String, sNumber2 As String, ByVal lBase As Long) As Boolean
    Dim lIndex As Long
    Dim lDigit1 As Long
    Dim lDigit2 As Long
    
    For lIndex = 1 To Len(sNumber1)
        lDigit1 = InStr(SYMBOLS, Mid$(sNumber1, lIndex, 1)) - 1
        lDigit2 = InStr(SYMBOLS, Mid$(sNumber2, lIndex, 1)) - 1
        If (lDigit1 + lDigit2) >= lBase Then
            CheckForCarry = True
            Exit Function
        End If
    Next
End Function

Public Function Convert(ByVal sNumber As String, ByVal iBase1 As Integer, ByVal iBase2 As Integer) As String
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

Public Function Pad(sNumber As String, iSize) As String
    If iSize > Len(sNumber) Then
        Pad = String$(iSize - Len(sNumber), "0") & sNumber
    Else
        Pad = Right$(sNumber, iSize)
    End If
End Function
