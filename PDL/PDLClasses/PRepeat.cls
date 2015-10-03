VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRepeat"
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
Private myUpperLimit As Long
Private myLowerLimit As Long
Private myNullControl As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    myResultControl = pcParseControl
    myNullControl = bNullControl
    
    myLowerLimit = 1
    
    Set myCondition = pInitParam(0)
    myConditionControl = pInitParam(1)
    
    If UBound(pInitParam) >= 2 Then
        myUpperLimit = pInitParam(2)
    End If
    If UBound(pInitParam) >= 3 Then
        myLowerLimit = pInitParam(3)
    End If
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim bFinish As Boolean
    Dim oResult As ParseTree
    Dim myPosition As Long
    Dim myCounter As Long
    Dim lIndexing As Long
    Dim myPreviousPosition As Long
    Dim lControl As ParseControl
    
    lIndexing = 1
    omyResult.Name = myName
    omyResult.Description = "REPEAT"
    
    myPosition = TextPosition
    While Not bFinish
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
                    If myUpperLimit = 0 And myCounter >= myLowerLimit Then
                        bFinish = True
                        IParseObject_Parse = True
                    End If
                End If
            Else
                bFinish = True
                If myCounter < myLowerLimit Then
                    omyResult.SubTree.Add oResult
                    IParseObject_Parse = False
                Else
                    IParseObject_Parse = True
                End If
            End If
        Else
            IParseObject_Parse = True
            bFinish = True
        End If
    Wend

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
    
    If IParseObject_Parse Then
        omyResult.Index = myCounter
        omyResult.Control = myResultControl
    Else
        omyResult.Index = myCounter
        omyResult.Error = True
        omyResult.TextStart = myPosition
    End If
End Function
