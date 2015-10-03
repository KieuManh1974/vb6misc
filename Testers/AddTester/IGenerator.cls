VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub GenerateItem()
' Interface
End Sub

Public Function GetItem() As Variant
' Interface
End Function

Public Function CheckAnswer(mvAnswer As Variant) As Boolean
' Interface
End Function

Public Function Equals(loItem As IGenerator) As Boolean
' Interface
End Function
