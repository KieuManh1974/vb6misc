VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POptional"
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
Private myCondition As IParseObject
Private myConditionControl As ParseControl
Private myCounter As Long

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    
    myResultControl = pcParseControl
    
    Set myCondition = pInitParam(0)
    myConditionControl = pInitParam(1)
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim bFinish As Boolean
    Dim myInnerPosition As Long
    Dim oResult As ParseTree
    Dim myPosition As Long
    Dim lIndexing As Long
    Dim lControl As ParseControl
    
    lIndexing = 1
    omyResult.Name = myName
    omyResult.Description = "OPTIONAL"
    
    myPosition = TextPosition
    
    Set oResult = New ParseTree
    oResult.Location = omyResult.Location & "(" & lIndexing & ")"
    If myCondition.Parse(oResult) Then
    
        lControl = myConditionControl
        If lControl = pcLeave Then
            lControl = oResult.Control
        End If
        
        If lControl <> pcOmit Then
            omyResult.Text = omyResult.Text & oResult.Text
            omyResult.TextStart = myPosition
            omyResult.TextEnd = TextPosition - 1
            
            If lControl = pcHide Then
                Set oResult.SubTree = Nothing
            End If
    
            lIndexing = lIndexing + 1
            omyResult.SubTree.Add oResult
        End If

        omyResult.Index = 1
    Else
        If myResultControl <> pcHide Then
            omyResult.Index = 0
        End If
    End If

    IParseObject_Parse = True
    omyResult.Control = myResultControl
End Function
