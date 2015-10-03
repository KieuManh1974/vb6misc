VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GenerateItem() As Variant
' Interface
End Function

Public Property Set Paper(oForm As Form)
'Interface
End Property

Public Sub Render(vItem As Variant)
'Interface
End Sub

Public Function Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
'Interface
End Function

Public Property Get SetSize() As Long
'Interface
End Property
