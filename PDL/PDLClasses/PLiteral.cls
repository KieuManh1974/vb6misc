VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PLiteral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IParseObject

' Generic variables

Private myResultControl As ParseControl
Private myName As String

' Object specific variables
Private myLiteralString As String
Private myOriginalLiteralString As String
Private myLiteralStringLength As Long
Private myCaseInsensitive As Boolean
Private myNullControl As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    
    myResultControl = pcParseControl
    
    myLiteralString = pInitParam(0)
    myLiteralStringLength = Len(myLiteralString)
    myOriginalLiteralString = myLiteralString
    If UBound(pInitParam) = 1 Then
        myCaseInsensitive = pInitParam(UBound(pInitParam))
        myLiteralString = UCase$(myLiteralString)
    End If
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim myPosition As Long
    Dim myGetChar As String
    Dim myOrigGetChar As String
    
    omyResult.Name = myName
    omyResult.Description = "LITERAL"
    
    myPosition = TextPosition
    If (TextPosition + myLiteralStringLength - 1) > LenTextString Then
        omyResult.TextStart = myPosition
        omyResult.Error = True
        Exit Function
    End If
    
    myOrigGetChar = Mid$(TextString, TextPosition, myLiteralStringLength)
        
    TextPosition = TextPosition + myLiteralStringLength
    
    myGetChar = myOrigGetChar
    If myCaseInsensitive Then
        myGetChar = UCase$(myGetChar)
    End If
    If myGetChar = myLiteralString Then
        omyResult.Control = myResultControl
        omyResult.Text = myOrigGetChar
        omyResult.TextStart = myPosition
        omyResult.TextEnd = TextPosition - 1
        IParseObject_Parse = True
    Else
        TextPosition = myPosition
        omyResult.TextStart = myPosition
        omyResult.Error = True
    End If
End Function
