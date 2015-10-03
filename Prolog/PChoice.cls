VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PChoice"
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

' // For choice collection
Private CanIgnore As Boolean


Private ErrorMessage As String

Private ParseCollection As New Collection



' // Initialise object as Choice
Public Sub Initialise(bIgnore As Boolean, bComplex As Boolean, bCanIgnore As Boolean, ParamArray vElements() As Variant)
    Dim iIndex As Long
    For iIndex = LBound(vElements) To UBound(vElements)
        ParseCollection.Add vElements(iIndex)
    Next
    Ignore = bIgnore
    Complex = bComplex
    CanIgnore = bCanIgnore
End Sub

Public Function IParseObject_FindString() As Collection
    Dim iIndex As Integer
    Dim ThisPosition As Long
    Dim vElement As IParseObject
    Dim ReturnCollection As Collection
    Dim ReturnString As String
        
    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition
    
    If Complex Then
        iIndex = 1
        For Each vElement In ParseCollection
            Set ReturnCollection = vElement.FindString()
            If ReturnCollection.Count <> 0 Then
                If Not IsNull(ReturnCollection(1)) Then
                    IParseObject_FindString.Add CStr(iIndex)
                    IParseObject_FindString.Add ReturnCollection
                End If
            Else
                IParseObject_FindString.Add CStr(iIndex)
                Exit For
            End If
            iIndex = iIndex + 1
        Next

        If ReturnCollection.Count <> 0 Then
            If IsNull(ReturnCollection(1)) Then
                If Not CanIgnore Then
                    'ErrorMessage = ErrorString
                    GoTo NotFound
                Else
                End If
            End If
        End If
    Else
        ReturnString = ""
        iIndex = 1
        For Each vElement In ParseCollection
            Set ReturnCollection = IParseObject_FindString(vElement)
            If ReturnCollection.Count <> 0 Then
                ReturnString = ReturnCollection(1)
                If Not IsNull(ReturnString) Then
                    IParseObject_FindString.Add ReturnString
                    Exit For
                End If
            Else
                Exit For
            End If
            iIndex = iIndex + 1
        Next

        If ReturnCollection.Count <> 0 Then
            If IsNull(ReturnCollection(1)) Then
                If Not CanIgnore Then
                    'ErrorMessage = .ErrorString
                    GoTo NotFound
                Else
                End If
            End If
        End If

    End If
    Exit Function
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function

