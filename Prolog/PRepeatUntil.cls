VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRepeatUntil"
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

' // For repeat until collection
Private TerminatorObject As IParseObject
Private ErrorMessage As String
Private ParseCollection As New Collection


'// Intialise object as RepeatUntil
Public Sub Initialise(bIgnore As Boolean, bComplex As Boolean, vTermCondition As Variant, ParamArray vElements() As Variant)

    Dim iIndex As Long
    For iIndex = LBound(vElements) To UBound(vElements)
        ParseCollection.Add vElements(iIndex)
    Next

    Complex = bComplex
    Ignore = bIgnore
    Set TerminatorObject = vTermCondition
End Sub

Public Function IParseObject_FindString() As Collection
    Dim ReturnString As String
    Dim bTerminatorFound As Boolean
    Dim lStringPosition As Long
    Dim iTempPosition As Integer
    Dim ReturnCollection As Collection
    Dim vElement As IParseObject
    Dim AString As String
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition

    If Complex Then
        While Not IsNull(ReturnString) And Not bTerminatorFound
            iTempPosition = lStringPosition
            Set ReturnCollection = TerminatorObject.FindString
            If ReturnCollection.Count = 0 Then
                lStringPosition = iTempPosition
                bTerminatorFound = True
            ElseIf Not IsNull(ReturnCollection(1)) Then
                lStringPosition = iTempPosition
                bTerminatorFound = True
            Else
                lStringPosition = iTempPosition
                For Each vElement In ParseCollection
                    Set ReturnCollection = IParseObject_FindString(vElement)
                    If ReturnCollection.Count <> 0 Then
                        If Not IsNull(ReturnCollection(1)) Then
                            IParseObject_FindString.Add ReturnCollection
                        Else
                            'ErrorMessage = .ErrorString
                            GoTo NotFound
                        End If
                    End If
                Next
            End If
        Wend

    Else
        While Not IsNull(ReturnString) And Not bTerminatorFound
            iTempPosition = lStringPosition
            Set ReturnCollection = TerminatorObject.FindString
            If ReturnCollection.Count = 0 Then
                lStringPosition = iTempPosition
                bTerminatorFound = True
            ElseIf Not IsNull(ReturnCollection(1)) Then
                lStringPosition = iTempPosition
                bTerminatorFound = True
            Else
                lStringPosition = iTempPosition
                For Each vElement In ParseCollection
                    Set ReturnCollection = vElement.FindString
                    If ReturnCollection.Count <> 0 Then
                        ReturnString = ReturnCollection(1)
                        If Not IsNull(ReturnString) Then
                            AString = AString & ReturnString
                        Else
                            'ErrorMessage = .ErrorString
                            GoTo NotFound
                        End If
                    End If
                Next
            End If
        Wend
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

