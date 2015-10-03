VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PIn"
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
Private myInString As String

Private myOriginalInString As String
Private myCaseInsensitive As Boolean
Private myNullControl As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    
    myResultControl = pcParseControl
    myInString = pInitParam(0)
    
    myOriginalInString = myInString
    If UBound(pInitParam) = 1 Then
        myCaseInsensitive = pInitParam(UBound(pInitParam))
        myInString = UCase$(myInString)
    End If
    
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim sTestChar As String
    Dim sTestGetChar As String
    Dim myPosition As Long
    
    InCount = InCount + 1
    
    omyResult.Name = myName
    omyResult.Description = "IN"
    
    myPosition = TextPosition
    If TextPosition > LenTextString Then
        omyResult.TextStart = myPosition
        omyResult.Error = True
        Exit Function
    End If
    
    sTestGetChar = Mid$(TextString, TextPosition, 1):
    TextPosition = TextPosition + 1
    
    If sTestGetChar <> "" Then
        sTestChar = sTestGetChar
        If myCaseInsensitive Then
            sTestChar = UCase$(sTestGetChar)
        End If
        
        If InStr(1, myInString, sTestChar, vbBinaryCompare) <> 0 Then
            omyResult.Control = myResultControl
            omyResult.Text = sTestGetChar
            omyResult.TextStart = myPosition
            omyResult.TextEnd = TextPosition - 1
            IParseObject_Parse = True
        Else
            TextPosition = myPosition
            omyResult.TextStart = myPosition
            omyResult.Error = True
        End If
    Else
        TextPosition = myPosition
        omyResult.TextStart = myPosition
        omyResult.Error = True
    End If
End Function
