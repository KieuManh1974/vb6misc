VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PSkip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IParseObject

' Generic variables

Private myResultControl As ParseControl
Private mySkipNo As Long
Private myName As String

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    
    mySkipNo = pInitParam(0)
    myResultControl = pcParseControl
    
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    IParseObject_Parse = True
    
    omyResult.Name = myName
    omyResult.Description = "SKIP"
    
    If (TextPosition + mySkipNo - 1) > LenTextString Then
        omyResult.Error = True
        omyResult.TextStart = TextPosition
        Exit Function
    End If

    omyResult.Text = Mid$(TextString, TextPosition, mySkipNo)
    omyResult.TextStart = TextPosition
    omyResult.TextEnd = TextPosition + mySkipNo - 1
    TextPosition = TextPosition + mySkipNo
    
    omyResult.Control = myResultControl
    IParseObject_Parse = True
End Function
