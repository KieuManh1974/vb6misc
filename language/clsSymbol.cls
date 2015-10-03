VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum stSymbolType
    stType
    stVar
End Enum

Public SymbolType As stSymbolType
Public Name As String
Public Unit As clsUnit
Public Repeater As clsUnit
Public AddressLocation As Long

Public Function ByteSize()
    ByteSize = Unit.ByteSize
End Function
