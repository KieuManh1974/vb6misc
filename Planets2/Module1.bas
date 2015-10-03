Attribute VB_Name = "Module1"
Option Explicit

Public Const UGA = 0.00000000006673
Public pi2 As Double

Public Function ASN(ByVal fSin As Double) As Double
    If fSin = 1 Then
        ASN = 2 * Atn(1)
    ElseIf fSin = -1 Then
        ASN = -2 * Atn(1)
    Else
        ASN = Atn(fSin / Sqr(1 - fSin * fSin))
    End If
End Function

Public Function ACS(ByVal fCos As Double) As Double
    ACS = 2 * Atn(1) - Atn(fCos / Sqr(1 - fCos * fCos))
End Function
