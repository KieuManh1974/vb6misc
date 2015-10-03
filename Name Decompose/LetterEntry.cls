VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LetterEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bNotEmptyArray As Boolean
Private Items() As LetterEntry

Public Score As Long
Public Letter As String

Public Property Get Count() As Long
    If bNotEmptyArray Then
        Count = UBound(Items) + 1
    End If
End Property

Public Property Get Item(lIndex As Long) As LetterEntry
    Set Item = Items(lIndex)
End Property


Public Function Add(sPart As String) As LetterEntry
    Dim ub As Long
    Dim lIndex As Long
    
    If bNotEmptyArray Then
        ub = UBound(Items)
        For lIndex = 0 To ub
            If Items(lIndex).Letter = sPart Then
                Items(lIndex).Score = Items(lIndex).Score + 1
                Set Add = Items(lIndex)
                Exit Function
            End If
        Next
        ReDim Preserve Items(ub + 1) As LetterEntry
        Set Items(ub + 1) = New LetterEntry
        Items(ub + 1).Letter = sPart
        Items(ub + 1).Score = 1
        Set Add = Items(ub + 1)
        
    Else
        ReDim Items(0) As LetterEntry
        Set Items(0) = New LetterEntry
        Items(0).Letter = sPart
        Items(0).Score = 1
        bNotEmptyArray = True
        Set Add = Items(0)
    End If
    

End Function
