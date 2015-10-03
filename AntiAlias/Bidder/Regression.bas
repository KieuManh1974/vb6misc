Attribute VB_Name = "Regression"
Option Explicit

' Linear regression
Public Function Slope(vValues1 As Variant, vValues2 As Variant) As Double
    Slope = (N(vValues1) * SumMult(vValues1, vValues2) - Sum(vValues1) * Sum(vValues2)) / (N(vValues1) * Sum2(vValues1) - Sum(vValues1) ^ 2)
End Function

Public Function Intercept(vValues1 As Variant, vValues2 As Variant) As Double
    Intercept = (Sum(vValues2) - Slope(vValues1, vValues2) * Sum(vValues1)) / N(vValues1)
End Function

Public Function RegLine(dSlope As Double, dIntercept As Double, dDatum As Double) As Double
    RegLine = dIntercept + dSlope * dDatum
End Function

' Exponential
Public Function EMultiplier(vValues1 As Variant, vValues2 As Variant) As Double
    EMultiplier = Exp(Intercept(vValues1, LogValues(vValues2)))
End Function

Public Function EBase(vValues1 As Variant, vValues2 As Variant) As Double
    EBase = Exp(Slope(vValues1, LogValues(vValues2)))
End Function

Public Function RegExp(dBase As Double, dMultiplier As Double, dDatum As Double) As Double
    RegExp = dMultiplier * dBase ^ dDatum
End Function

Public Function LogValues(vValues As Variant) As Variant
    Dim vValue As Variant
    Dim vLogValues As Variant
    
    vLogValues = Array()
    
    For Each vValue In vValues
        ReDim Preserve vLogValues(UBound(vLogValues) + 1) As Variant
        vLogValues(UBound(vLogValues)) = Log(vValue)
    Next
    LogValues = vLogValues
End Function

Public Function Cov(vValues1 As Variant, vValues2 As Variant) As Double
    Cov = SCP(vValues1, vValues2) / (N(vValues1) - 1)
End Function

Public Function SCP(vValues1 As Variant, vValues2 As Variant) As Double
    SCP = SumMult(vValues1, vValues2) - (Sum(vValues1) * Sum(vValues2)) / N(vValues1)
End Function

Public Function SD(vValues As Variant) As Double
    SD = Sqr(SS(vValues) / (N(vValues) - 1))
End Function

Public Function SS(vValues As Variant) As Double
    SS = Sum2(vValues) - Sum(vValues) ^ 2 / N(vValues)
End Function

Public Function Sum(vValues As Variant) As Double
    Dim vValue As Variant
    
    For Each vValue In vValues
        Sum = Sum + CDbl(vValue)
    Next
End Function


Public Function Sum2(vValues As Variant) As Double
    Dim vValue As Variant
    
    For Each vValue In vValues
        Sum2 = Sum2 + CDbl(vValue) * CDbl(vValue)
    Next
End Function

Public Function SumMult(vValues1 As Variant, vValues2 As Variant) As Double
    Dim lIndex As Long
    
    For lIndex = LBound(vValues1) To UBound(vValues1)
        SumMult = SumMult + vValues1(lIndex) * vValues2(lIndex)
    Next
End Function


Public Function N(vValues As Variant) As Double
    N = UBound(vValues) - LBound(vValues) + 1
End Function

' Attempt to fit to curve y=B(1-exp(A*-x))
Public Function NumericalRegression(vValuesX As Variant, vValuesY As Variant, dOffset As Double) As Variant
    Dim bFinished As Boolean
    Dim dFitness As Double
    Dim dMaxFitness As Double
    Dim dDelta As Double
    Dim lNoVariables As Long
    Dim lDeltas() As Long
    Dim lOptimums() As Long
    Dim lVariableIndex As Long
    Dim dVariables() As Double
    Dim bOptimum As Boolean
    Dim vVariables As Variant
       
    lNoVariables = 2
    ReDim lDeltas(lNoVariables - 1)
    ReDim lOptimums(lNoVariables - 1)
    ReDim dVariables(lNoVariables - 1)
    
    For lVariableIndex = 0 To lNoVariables - 1
        lDeltas(lVariableIndex) = -1
    Next
    
    dVariables(0) = 0
    dVariables(1) = 0
    'dVariables(2) = -dOffset
    dDelta = 8192
    On Error Resume Next: dMaxFitness = 1 / 0: On Error GoTo 0

    While Not bFinished
        bOptimum = True
        Do
            dFitness = FitnessLinear(vValuesX, vValuesY, dVariables(0) + CDbl(lDeltas(0)) * dDelta, dVariables(1) + CDbl(lDeltas(1)) * dDelta)
            'Debug.Print dFitness
            If dFitness < dMaxFitness Then
                bOptimum = False
                'DoEvents
                dMaxFitness = dFitness
                
                For lVariableIndex = 0 To lNoVariables - 1
                    lOptimums(lVariableIndex) = lDeltas(lVariableIndex)
                Next
            End If
            
            lVariableIndex = 0
            Do
                lDeltas(lVariableIndex) = lDeltas(lVariableIndex) + 1
                If lDeltas(lVariableIndex) = 2 Then
                    lDeltas(lVariableIndex) = -1
                    lVariableIndex = lVariableIndex + 1
                Else
                    Exit Do
                End If
            Loop Until lVariableIndex = lNoVariables
        Loop Until lVariableIndex = lNoVariables
        
        
        If bOptimum Then
            dDelta = dDelta / 2
            If dDelta = 0 Then
                bFinished = True
            End If
        Else
            For lVariableIndex = 0 To lNoVariables - 1
                'Debug.Print lOptimums(lVariableIndex);
                dVariables(lVariableIndex) = dVariables(lVariableIndex) + CDbl(lOptimums(lVariableIndex)) * dDelta
            Next
            'Debug.Print
        End If
    Wend

    vVariables = Array()
    ReDim vVariables(lNoVariables - 1)
    
    For lVariableIndex = 0 To lNoVariables - 1
        vVariables(lVariableIndex) = dVariables(lVariableIndex)
    Next
        
    NumericalRegression = vVariables
End Function

'Private Function Fitness(vValuesX As Variant, vValuesY As Variant, dA As Double, dB As Double) As Double
'    Dim lIndex As Long
'
'    For lIndex = LBound(vValuesX) To UBound(vValuesX)
'        Fitness = Fitness + (vValuesY(lIndex) - (dB * (1 - Exp(-dA * vValuesX(lIndex))))) ^ 2
'    Next
'End Function

'' exponential a*b^t
'Private Function Fitness(vValuesX As Variant, vValuesY As Variant, dA As Double, dB As Double) As Double
'    Dim lIndex As Long
'
'    For lIndex = LBound(vValuesX) To UBound(vValuesX)
'        Fitness = Fitness + (vValuesY(lIndex) - (dA * dB ^ vValuesX(lIndex))) ^ 2
'    Next
'End Function

' quadratic  a*x^2+b*x+0
Private Function FitnessQuadratic(vValuesX As Variant, vValuesY As Variant, dA As Double, dB As Double, dC As Double) As Double
    Dim lIndex As Long
    Dim dDiff As Double
    
    On Error GoTo FitnessError
    
    For lIndex = LBound(vValuesX) To UBound(vValuesX)
        dDiff = vValuesY(lIndex) - (dA * (vValuesX(lIndex) + dC) ^ 2 + dB * (vValuesX(lIndex) + dC))
        Fitness = Fitness + dDiff * dDiff
    Next
    
    Exit Function
FitnessError:
    On Error Resume Next: Fitness = 1 / 0: On Error GoTo 0
End Function

' linear y= a*x+b
Private Function FitnessLinear(vValuesX As Variant, vValuesY As Variant, dA As Double, dB As Double) As Double
    Dim lIndex As Long
    Dim dDiff As Double
    
    On Error GoTo FitnessError
    
    For lIndex = LBound(vValuesX) To UBound(vValuesX)
        dDiff = vValuesY(lIndex) - (dA * vValuesX(lIndex) + dB)
        FitnessLinear = FitnessLinear + dDiff * dDiff
    Next
    
    Exit Function
FitnessError:
    On Error Resume Next: FitnessLinear = 1 / 0: On Error GoTo 0
End Function

