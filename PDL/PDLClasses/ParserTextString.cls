VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParserTextString"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ParseControl
    pcNormal = 0
    pcIgnore = 1
    pcOmit = 2
End Enum

Public Property Let ParserText(sTextString As String)
    TextString = sTextString
    TextPosition = 1
    LenTextString = Len(sTextString)
End Property

Public Property Get ParserText() As String
    ParserText = TextString
End Property

Public Sub ResetText()
    TextPosition = 1
End Sub

Public Property Let ParserTextPosition(lPosition As Long)
    TextPosition = lPosition
End Property

Public Property Get ParserTextPosition() As Long
    ParserTextPosition = TextPosition
End Property

Public Property Set External(oExternal As Object)
    Set ExternalClass = oExternal
End Property
