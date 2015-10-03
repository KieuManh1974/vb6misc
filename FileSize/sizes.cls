VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SizeInfo
    ByteSize As Long
    Occurancce As Long
End Type

Private bSorted As Boolean

Private SizeArray() As SizeInfo

Public Sub Include(lByteSize)
    Dim ub As Long
    Dim x As Long
    ub = UBound(SizeArray)
    
    For x = 0 To ub
        If SizeArray(x).ByteSize = lByteSize Then
            SizeArray(x).Occurancce = SizeArray(x).Occurancce + 1
            Exit Sub
        End If
    Next
    
    ReDim Preserve SizeArray(ub + 1) As SizeInfo
    
    SizeArray(ub + 1).ByteSize = lByteSize
    SizeArray(ub + 1).Occurancce = 1
End Sub

Public Function GetItemSize(lItem As Long) As Long
    If Not bSorted Then
        SortArray
    End If
    GetItemSize = SizeArray(lItem).ByteSize
End Function

Public Function GetItemOccurance(lItem As Long) As Long
    If Not bSorted Then
        SortArray
    End If
    GetItemOccurance = SizeArray(lItem).Occurancce
End Function

Public Property Get MaxArray() As Long
    MaxArray = UBound(SizeArray)
End Property

Private Sub SortArray()
    Dim bFinished As Boolean
    Dim ub As Long
    Dim temp As SizeInfo
    Dim x As Long
    
    ub = UBound(SizeArray) - 1
    Do
        bFinished = True
        For x = 0 To ub
            If SizeArray(x + 1).ByteSize > SizeArray(x).ByteSize Then
                temp = SizeArray(x)
                SizeArray(x) = SizeArray(x + 1)
                SizeArray(x + 1) = temp
                bFinished = False
            End If
        Next
    Loop While Not bFinished
    bSorted = True
End Sub

Private Sub Class_Initialize()
    ReDim SizeArray(0) As SizeInfo
End Sub
