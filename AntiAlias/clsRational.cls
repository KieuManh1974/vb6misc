VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRational"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Numerator As Long
Public Denominator As Long

Public Property Get Value() As Double
    Value = Numerator / CDbl(Denominator)
End Property

Public Function Add(oRational As clsRational) As clsRational
    Set Add = New clsRational
    
    Add.Numerator = Me.Numerator * oRational.Denominator + Me.Denominator * oRational.Numerator
    Add.Denominator = Me.Denominator * oRational.Denominator
End Function

Public Function Subtract(oRational As clsRational) As clsRational
    Set Subtract = New clsRational
    
    Subtract.Numerator = Me.Numerator * oRational.Denominator - Me.Denominator * oRational.Numerator
    Subtract.Denominator = Me.Denominator * oRational.Denominator

End Function

Public Function Multiply(oRational As clsRational) As clsRational
    Set Multiply = New clsRational
    
    Multiply.Numerator = Me.Numerator * oRational.Numerator
    Multiply.Denominator = Me.Denominator * oRational.Denominator
End Function

Public Function Divide(oRational As clsRational) As clsRational
    Set Divide = New clsRational
    
    Divide.Numerator = Me.Numerator * oRational.Denominator
    Divide.Denominator = Me.Denominator * oRational.Numerator

End Function

Public Function Reduce()
    Dim lValue1 As Long
    Dim lValue2 As Long
    Dim lTemp As Long
    
    lValue1 = Numerator
    lValue2 = Denominator
    
    While lValue2 <> 0
        lTemp = lValue2
        lValue2 = lValue1 Mod lValue2
        lValue1 = lTemp
    Wend
    
    Numerator = Numerator / lValue1
    Denominator = Denominator / lValue1
End Function
