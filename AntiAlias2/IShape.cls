VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get Red() As Long
'
End Property

Public Property Get Green() As Long
'
End Property

Public Property Get Blue() As Long
'
End Property

Public Sub StartScan()
'
End Sub

Public Function Inside() As Boolean
'
End Function

Public Property Let Colour(ByVal lColour As Long)
'
End Property

Public Property Get Colour() As Long
'
End Property

Public Property Let Inverse(ByVal bInverse As Boolean)
'
End Property


Public Sub NextScanline()
'
End Sub

Public Sub NextPixel()
'
End Sub

Public Sub NextSubScanline()
'
End Sub

Public Sub NextSubPixel()
'
End Sub
