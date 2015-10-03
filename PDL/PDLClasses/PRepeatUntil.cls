VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRepeatUntil"
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
Private myTerminator As IParseObject
Private myLowerLimit As Long
Private myUpperLimit As Long
Private myNullControl As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    myResultControl = pcParseControl
    myNullControl = bNullControl
    
    myLowerLimit = 1
    
    Set myCondition = pInitParam(0)
    myConditionControl = pInitParam(1)
    Set myTerminator = pInitParam(2)
    
    If UBound(pInitParam) >= 3 Then
        myUpperLimit = pInitParam(3)
    End If
    If UBound(pInitParam) >= 4 Then
        myLowerLimit = pInitParam(4)
    End If
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim bFinish As Boolean
    Dim myInnerPosition As Long
    Dim oResult As ParseTree
    Dim myPosition As Long
    Dim myCounter As Long
    Dim lIndexing As Long
    Dim myPreviousPosition As Long
    Dim lControl As ParseControl
    Dim bByTerminator As Boolean
    
    If myName = "text" Then
        Stop
    End If
    
    lIndexing = 1
    omyResult.Name = myName
    omyResult.Description = "REPEATUNTIL"
    
    myPosition = TextPosition
    While Not bFinish
        myInnerPosition = TextPosition
        Set oResult = New ParseTree
        oResult.Location = omyResult.Location & "(" & lIndexing & ")"
        If myTerminator.Parse(oResult) Then
            TextPosition = myInnerPosition
            bFinish = True
            If myCounter < myLowerLimit Then
                omyResult.Control = myResultControl
                Exit Function
            End If
            omyResult.Index = myCounter
            bByTerminator = True
            IParseObject_Parse = True
        Else
            TextPosition = myInnerPosition
            If myCounter < myUpperLimit Or myUpperLimit = 0 Then
                Set oResult = New ParseTree
                oResult.Location = omyResult.Location & "(" & lIndexing & ")"
                myPreviousPosition = TextPosition
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
                    myCounter = myCounter + 1
                    If TextPosition = myPreviousPosition Then
                        bFinish = True
                    End If
                Else
                    TextPosition = myPosition
                    omyResult.Index = myCounter
                    omyResult.SubTree.Add oResult
                    IParseObject_Parse = False
                    bFinish = True
                End If
            Else
                IParseObject_Parse = True
                bFinish = True
            End If
        End If
    Wend

    If myNullControl Then
        If IParseObject_Parse = True And Not bByTerminator Then
            If myPosition = TextPosition Then
                IParseObject_Parse = False
                omyResult.Error = True
                omyResult.TextStart = myPosition
                Exit Function
            End If
        End If
    End If
    
    If Not IParseObject_Parse Then
        omyResult.Error = True
        omyResult.TextStart = myPosition
    Else
        omyResult.Control = myResultControl
    End If
End Function


