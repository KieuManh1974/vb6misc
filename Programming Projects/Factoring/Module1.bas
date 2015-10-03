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
    
    'Debug.Print oSquare.Display
    For lNumber = 256 To 300
        If Not Prime(lNumber) Then
            Set oSquare = New clsSquare
            oSquare.Size = 10
            oSquare.Initialise lNumber
            For lRepeat = 1 To 100000
                nTime = 0
                StartCounter
                While lScore > 0 Or oSquare.Factor(0) = 1 Or oSquare.Factor(1) = 1
                    Set oClone1 = oSquare.Clone
                    Set oClone2 = oSquare.Clone
                    oClone1.Mutate
                    oClone2.Mutate
                    
                    If oClone1.MismatchCount < oClone2.MismatchCount Then
                        Set oSquare = oClone1
                    Else
                        Set oSquare = oClone2
                    End If
                    
                    lScore = oSquare.MismatchCount
                    
            '        lScore = oClone.MismatchCount
            '        If lScore <= lPreviousScore Then
            '            Set oSquare = oClone
            '            lPreviousScore = lScore
            '            Debug.Print oSquare.Display
            '            If 1 = 1 Then
            '            End If
            '        End If
                    
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
