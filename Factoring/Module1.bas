Attribute VB_Name = "Module1"
Option Explicit

Private mlSquare(7, 7) As Long
Private mlFactor(1, 7) As Long

Public Sub main()
    Factor
End Sub

Public Sub Factor()
    Dim oSquare As New clsSquare
    Dim oClone1 As clsSquare
    Dim oClone2 As clsSquare
    Dim lPreviousScore As Long
    Dim lScore As Long
    Dim bFinished As Boolean
    Dim nTime As Single
    Dim nMaxTime As Single
    Dim lRepeat As Long
    Dim lNumber As Long
    Dim lScore1 As Long
    Dim lScore2 As Long
    
    Dim lMix As Long
    
    'Debug.Print oSquare.Display
    For lNumber = 200 To 300
        DoEvents
        If Not Prime(lNumber) Then
            nTime = 0
            For lRepeat = 1 To 1
                Set oSquare = New clsSquare
                oSquare.Initialise lNumber
                
                'Debug.Print oSquare.Display
                'Debug.Print oSquare.Sum
                For lMix = 1 To 200
                    oSquare.Mutate2
                Next
                'Debug.Print oSquare.Display
               ' Debug.Print oSquare.Sum
                
                'Debug.Print oSquare.Factor(0) * oSquare.Factor(1)
            
                StartCounter
                lScore = oSquare.MismatchCount
                
                StartCounter 1
                
                While (lScore > 0 Or oSquare.Factor(0) = 1 Or oSquare.Factor(1) = 1)
                    Set oClone1 = oSquare.Clone
                    Set oClone2 = oSquare.Clone
                    oClone1.Mutate
                    oClone2.Mutate
                    
                    lScore1 = oClone1.MismatchCount
                    lScore2 = oClone2.MismatchCount
                    
                    If lScore1 < lScore2 Then
                        Set oSquare = oClone1
                        lScore = lScore1
                    Else
                        Set oSquare = oClone2
                        lScore = lScore2
                    End If
'                    If GetCounter(1) > 2.5 Then
'                        For lMix = 1 To 100
'                            oSquare.Mutate
'                        Next
'                        lScore = oSquare.MismatchCount
'                        StartCounter 1
'                    End If
                Wend
                nTime = nTime + GetCounter
            Next
            Debug.Print lNumber & "  " & Format$(nTime / (lRepeat - 1), "0.000000000000") & "  "; oSquare.Factor(0) & " x " & oSquare.Factor(1)
        End If
    Next
End Sub

Public Function Prime(ByVal lNumber As Long) As Boolean
    Dim lDivisor As Long
    Dim lDivisorTest As Long
    
    Prime = True
    For lDivisor = 1 To Sqr(lNumber) Step 2
        If lDivisor = 1 Then
            lDivisorTest = 2
        Else
            lDivisorTest = lDivisor
        End If
        If (lNumber Mod lDivisorTest) = 0 Then
            Prime = False
            Exit Function
        End If
    Next
End Function
