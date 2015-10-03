VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCharacterRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParseObject

' // Result ignore
Private Ignore As Boolean

' // For character range
Private StartRange As String ' * 1
Private EndRange As String ' * 1

Private ErrorMessage As String
Private ParseCollection As New Collection



' // Intialise object as CharacterRange
Public Sub Initialise(bIgnore As Boolean, StartOfRange As String, EndOfRange As String)
    StartRange = StartOfRange
    EndRange = EndOfRange
    Ignore = bIgnore
End Sub

Public Function IParseObject_FindString() As Collection
    Dim ThisChar As String * 1
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    
    ThisChar = GetChar
    If ThisChar < StartRange Or ThisChar > EndRange Then
        'ErrorMessage = ErrorString
        GoTo NotFound
    End If

    If Not Ignore Then
        IParseObject_FindString.Add CStr(ThisChar)
    End If
    Exit Function
    
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function


