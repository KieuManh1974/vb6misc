VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Device"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlhDC(1) As Long
Private mlhWnd(1) As Long
Private mlLeftTableOffset(1) As Long
Private mlTopTableOffset(1) As Long
Public moDevice As PictureBox

Public Property Get hDC(ByVal lIndex As Long) As Long
    hDC = mlhDC(lIndex)
End Property

Public Property Get hWnd(ByVal lIndex As Long) As Long
    hWnd = mlhWnd(lIndex)
End Property

Public Property Let hDC(ByVal lIndex As Long, ByVal lValue As Long)
    mlhDC(lIndex) = lValue
End Property

Public Property Let hWnd(ByVal lIndex As Long, ByVal lValue As Long)
    mlhWnd(lIndex) = lValue
End Property


Public Property Get LeftTableOffset(ByVal lIndex As Long) As Long
    LeftTableOffset = mlLeftTableOffset(lIndex)
End Property

Public Property Get TopTableOffset(ByVal lIndex As Long) As Long
    TopTableOffset = mlTopTableOffset(lIndex)
End Property

Public Property Let LeftTableOffset(ByVal lIndex As Long, ByVal lValue As Long)
    mlLeftTableOffset(lIndex) = lValue
End Property

Public Property Let TopTableOffset(ByVal lIndex As Long, ByVal lValue As Long)
    mlTopTableOffset(lIndex) = lValue
End Property
