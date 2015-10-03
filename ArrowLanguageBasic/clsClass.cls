VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Identifier As String
Public Range As New clsRange

Public Property Get Size() As Long
    Dim lTotalRange As Long
    Dim lBytes As Long
    
    lTotalRange = Range.Ending - Range.Starting
    
    If lTotalRange > 0 Then
        Size = Int(Log(lTotalRange) / Log(256) + 1)
    Else
        Size = Int(Log(Range.Starting) / Log(256) + 1)
    End If
End Property

