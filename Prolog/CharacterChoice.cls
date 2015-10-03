VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCharacterChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParseObject

' // Result ignore
Private Ignore As Boolean

' // For character choice
Private CharChoice As String

Private ErrorMessage As String



' // Initialise object as CharacterChoice
Public Sub Initialise(bIgnore As Boolean, sChars As String)
    CharChoice = sChars
    Ignore = bIgnore
End Sub

Public Function IParseObject_FindString() As Collection
    Dim ThisChar As String * 1
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition

    ThisChar = GetChar
    If InStr(CharChoice, ThisChar) = 0 Then
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


