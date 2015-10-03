VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParseObject

' // Result ignore
Private Ignore As Boolean

' // For join, choice, repeat or repeat until collection
Private Complex As Boolean


Private ErrorMessage As String

Private ParseCollection As New Collection



' // Initialise object as Join
Public Sub Initialise(bIgnore As Boolean, bComplex As Boolean, ParamArray vElements() As Variant)
    Dim iIndex As Long
    For iIndex = LBound(vElements) To UBound(vElements)
        ParseCollection.Add vElements(iIndex)
    Next
    
    Ignore = bIgnore
    Complex = bComplex
End Sub

Public Function IParseObject_FindString() As Collection
    Dim vElement As IParseObject
    Dim ReturnCollection As New Collection
    Dim AString As String
    Dim ReturnString As String
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection

    If Complex Then
        For Each vElement In ParseCollection
            Set ReturnCollection = vElement.FindString(vElement)
            If ReturnCollection.Count <> 0 Then
                If IsNull(ReturnCollection(1)) Then
                    GoTo NotFound
                End If
                IParseObject_FindString.Add ReturnCollection
            End If
        Next
    Else
        AString = ""
        For Each vElement In ParseCollection
            Set ReturnCollection = IParseObject_FindString(vElement)
            If ReturnCollection.Count <> 0 Then
                ReturnString = vElement.FindString()(1)
                If IsNull(ReturnCollection(1)) Then
                    'ErrorMessage = .ErrorString
                    GoTo NotFound
                End If
                AString = AString & ReturnCollection(1)
            End If
        Next
        IParseObject_FindString.Add AString
    End If
    Exit Function
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function

