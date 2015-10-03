VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rational"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Numerator As Long
Public Denominator As Long

Public Property Get Value(Optional lLength As Long) As String

End Property

Public Sub Assign(lNumerator As Long, lDenominator As Long)
    Numerator = lNumerator
    Denominator = lDenominator
End Sub

Public Function Add(oRational As Rational)
    Dim lSum1 As Long
    Dim lSum2 As Long
    
    lSum1 = Numerator * oRational.Denominator
    lSum2 = oRational.Numerator * Denominator
    
    Numerator = lSum1 + lSum2
    Denominator = Denominator * oRational.Denominator
    Reduce
End Function

Public Function Subtract(oRational As Rational)
    Dim lSum1 As Long
    Dim lSum2 As Long
    
    lSum1 = Numerator * oRational.Denominator
    lSum2 = oRational.Numerator * Denominator
    
    Numerator = lSum1 - lSum2
    Denominator = Denominator * oRational.Denominator
    Reduce
End Function

Public Function Multiply(oRational As Rational)
    Numerator = Numerator * oRational.Numerator
    Denominator = Denominator * oRational.Denominator
    Reduce
End Function

Public Function Divide(oRational As Rational)
    Numerator = Numerator * oRational.Denominator
    Denominator = Denominator * oRational.Numerator
    Reduce
End Function

Private Sub Reduce()
    Dim lHCF As Long
    
    If Numerator < 0 Then
        Numerator = -Numerator
        Denominator = -Denominator
    End If
    lHCF = Euclidian(Abs(Numerator), Abs(Denominator))
    Numerator = Numerator / lHCF
    Denominator = Denominator / lHCF
End Sub

Private Function Euclidian(ByVal lFactor1 As Long, ByVal lFactor2 As Long) As Long
    Dim t As Long
    Dim r As Long
    
    Do
        t = Int(lFactor1 / lFactor2)
        r = lFactor1 - t * lFactor2
        lFactor1 = lFactor2
        lFactor2 = r
    Loop Until r = 0
    Euclidian = lFactor1
End Function
