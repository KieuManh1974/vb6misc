VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PAnd"
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
Private myObjects() As IParseObject
Private myResultControls() As ParseControl
Private myObjectCount As Long
Private myNullControl As Boolean

Private myReplace As Boolean

' Initialises parameters used for parsing
Public Function IParseObject_Initialise(sName As String, pcParseControl As ParseControl, bNullControl As Boolean, ParamArray pInitParam() As Variant) As IParseObject
    myName = sName
    myResultControl = pcParseControl
    myNullControl = bNullControl

    Dim lObjectIndex As Long
    Dim lAndObjectIndex As Long

    If TypeOf pInitParam(0) Is Collection Then
        myObjectCount = pInitParam(0).Count \ 2
        ReDim myObjects(1 To myObjectCount) As IParseObject
        ReDim myResultControls(1 To myObjectCount) As ParseControl
        For lObjectIndex = 1 To myObjectCount
            Set myObjects(lObjectIndex) = pInitParam(0)(lObjectIndex * 2 - 1)
            myResultControls(lObjectIndex) = pInitParam(0)(lObjectIndex * 2)
        Next
    Else
        myObjectCount = (UBound(pInitParam) - LBound(pInitParam) + 1) \ 2
        ReDim myObjects(1 To myObjectCount) As IParseObject
        ReDim myResultControls(1 To myObjectCount) As ParseControl
        lObjectIndex = 1
        For lAndObjectIndex = LBound(pInitParam) To UBound(pInitParam) Step 2
            Set myObjects(lObjectIndex) = pInitParam(lAndObjectIndex)
            myResultControls(lObjectIndex) = pInitParam(lAndObjectIndex + 1)
            lObjectIndex = lObjectIndex + 1
        Next
    End If

    Set IParseObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function IParseObject_Parse(omyResult As ParseTree) As Boolean
    Dim oAndObject As IParseObject
    Dim oResult As ParseTree
    Dim myPosition As Long
    Dim lObjectIndex As Long
    Dim lIndexing As Long
    Dim lControl As ParseControl
    
    lIndexing = 1
    omyResult.Name = myName
    omyResult.Description = "AND"

    myPosition = TextPosition
    IParseObject_Parse = True
    For lObjectIndex = 1 To myObjectCount
        Set oResult = New ParseTree
        oResult.Location = omyResult.Location & "(" & lIndexing & ")"
        If Not myObjects(lObjectIndex).Parse(oResult) Then
            IParseObject_Parse = False
            TextPosition = myPosition
            omyResult.Error = True
            omyResult.TextStart = myPosition
            omyResult.Index = lObjectIndex
            lIndexing = lIndexing + 1
            omyResult.SubTree.Add oResult
            Exit For
        End If

        lControl = myResultControls(lObjectIndex)
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
    Next

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

    omyResult.Control = myResultControl
End Function
