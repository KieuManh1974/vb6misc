VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Ranges() As clsRange
Public RangesCount As Long
Public MaxValue As Long
Public Size As Long

Public Sub AddRange(oRange As clsRange)
    ReDim Ranges(RangesCount)
    Set Ranges(RangesCount) = oRange
    RangesCount = RangesCount + 1
    Size = Size + oRange.Size
End Sub
