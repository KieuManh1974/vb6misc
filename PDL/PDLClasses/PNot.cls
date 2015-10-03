VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IParseObject

Private myName As String
Private myResultControl As ParseControl
Private myNotObject As IParseObject

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    
    myResultControl = pcParseControl
    Set myNotObject = pInitParam(LBound(pInitParam))
    
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim oResult As New ParseTree
    
    omyResult.Name = myName
    omyResult.Description = "NOT"
    omyResult.TextStart = TextPosition
    omyResult.Control = myResultControl
    
    IParseObject_Parse = Not myNotObject.Parse(oResult)
    If Not IParseObject_Parse Then
        omyResult.Error = True
    End If
End Function
