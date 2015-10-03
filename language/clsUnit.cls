VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Fields() As clsField
Public FieldsCount As Long
Public Size As Long ' bytes

Public Sub AddField(oField As clsField)
    ReDim Fields(FieldsCount)
    Set Fields(FieldsCount) = oField
    FieldsCount = FieldsCount + 1
    Size = Size * oField.Size
End Sub

Public Function ByteSize() As Long
    Dim dValue As Single
    
    dValue = Size - 1
    While dValue >= 1
        dValue = dValue / 256
        ByteSize = ByteSize + 1
    Wend
End Function

Private Sub Class_Initialize()
    Size = 1
End Sub
