VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PExternal"
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

Private myFunction As String
Private myParameters As String

Private myNullControl As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    myResultControl = pcParseControl
    myNullControl = bNullControl

    myFunction = pInitParam(0)
    myParameters = pInitParam(1)
    
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim myPosition As Long
    Dim vParameters() As Variant
    Dim vFunctionParameters As Variant
    Dim iParameterIndex As Long
    
    Dim iIndex As Integer
    Dim sText As String
    
    ExternalCount = ExternalCount + 1
    
    omyResult.Name = myName
    omyResult.Description = "EXTERNAL"
    
    vFunctionParameters = Split(myParameters, ",", , vbTextCompare)
    
    myPosition = TextPosition
    
    If CallByName(ExternalClass, myFunction, VbMethod, TextString, TextPosition, iIndex, sText, vFunctionParameters) Then
        omyResult.Control = myResultControl
        If myResultControl <> pcOmit Then
            omyResult.Text = sText
            omyResult.TextStart = myPosition
            omyResult.TextEnd = TextPosition - 1
        End If
        
        IParseObject_Parse = True
        If myNullControl Then
            If IParseObject_Parse = True Then
                If myPosition = TextPosition Then
                    IParseObject_Parse = False
                    omyResult.Error = True
                    omyResult.TextStart = myPosition
                    Exit Function
                End If
            End If
        End If
    Else
        TextPosition = myPosition
        omyResult.TextStart = myPosition
        omyResult.Error = True
    End If
End Function

