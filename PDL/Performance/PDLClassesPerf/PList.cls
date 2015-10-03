VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PList"
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
Private myDelimiter As IParseObject
Private myDelimiterControl As ParseControl
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
    Set myDelimiter = pInitParam(2)
    myDelimiterControl = pInitParam(3)
    
    If UBound(pInitParam) >= 4 Then
        myUpperLimit = pInitParam(4)
    End If
    If UBound(pInitParam) >= 5 Then
        myLowerLimit = pInitParam(5)
    End If
    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim bFinish As Boolean
    Dim oResult As ParseTree
    Dim oDelimiterResult As ParseTree
    Dim bAddDelimiter As Boolean
    Dim myPosition As Long
    Dim myCounter As Long
    Dim lIndexing As Long
    Dim myInitialPosition As Long
    Dim lControl As ParseControl
    
    ListCount = ListCount + 1
        
    lIndexing = 1
    omyResult.Name = myName
    omyResult.Description = "LIST"
    
    myPosition = TextPosition
    myInitialPosition = myPosition
    
    ' First member of list
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
        myCounter = myCounter + 1
    Else
        If myLowerLimit <> 0 Then
            omyResult.Index = myCounter
            omyResult.SubTree.Add oResult
            IParseObject_Parse = False
            bFinish = True
        Else
            IParseObject_Parse = True
            bFinish = True
        End If
    End If
    
    While Not bFinish
        If myCounter < myUpperLimit Or myUpperLimit = 0 Then
            Set oDelimiterResult = New ParseTree
            oDelimiterResult.Location = omyResult.Location & "(" & lIndexing & ")"
            myPosition = TextPosition
            If myDelimiter.Parse(oDelimiterResult) Then
                Set oResult = New ParseTree
                oResult.Location = omyResult.Location & "(" & lIndexing & ")"
                If myCondition.Parse(oResult) Then
                    lControl = myDelimiterControl
                    If lControl = pcLeave Then
                        lControl = oDelimiterResult.Control
                    End If
                    
                    If lControl <> pcOmit Then
                        omyResult.Text = omyResult.Text & oDelimiterResult.Text
                        omyResult.TextEnd = TextPosition - 1
                        
                        If lControl = pcHide Then
                            Set oDelimiterResult.SubTree = Nothing
                        End If
                
                        lIndexing = lIndexing + 1
                        omyResult.SubTree.Add oDelimiterResult
                    End If
                
                    lControl = myConditionControl
                    If lControl = pcLeave Then
                        lControl = oResult.Control
                    End If
                    
                    If lControl <> pcOmit Then
                        omyResult.Text = omyResult.Text & oResult.Text
                        omyResult.TextEnd = TextPosition - 1
                        
                        If lControl = pcHide Then
                            Set oResult.SubTree = Nothing
                        End If
                
                        lIndexing = lIndexing + 1
                        omyResult.SubTree.Add oResult
                    End If
                    myCounter = myCounter + 1
                Else
                    TextPosition = myPosition
                    bFinish = True
                    If myCounter < myLowerLimit Then
                        omyResult.Index = myCounter
                        omyResult.SubTree.Add oResult
                        IParseObject_Parse = False
                    Else
                        IParseObject_Parse = True
                    End If
                End If
            Else
                bFinish = True
                If myCounter < myLowerLimit Then
                    omyResult.Index = myCounter
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
        omyResult.Error = True
        omyResult.TextStart = TextPosition
    End If
End Function


