VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Identifier As String
Public UnitClass As clsClass
Public IsConstant As Boolean

Private mlAddress As Long

Public Property Get MyAddress(ByVal lAddress As Long)
    If mlAddress = -1 Then
        mlAddress = lAddress
    End If
    MyAddress = mlAddress
End Property

Private Sub Class_Initialize()
    mlAddress = -1
End Sub
