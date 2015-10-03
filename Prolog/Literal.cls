VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLiteral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParseObject

' // Result ignore
Private Ignore As Boolean

' // For literal
Private LiteralString As String

Private ErrorMessage As String



' // Initialise object as Literal
Public Sub Initialise(bIgnore As Boolean, sLiteral As String)
    Ignore = bIgnore
    LiteralString = sLiteral
End Sub

Public Function IParseObject_FindString() As Collection
    Dim iIndex As Integer
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    
    For iIndex = 1 To Len(LiteralString)
        If GetChar <> Mid$(LiteralString, iIndex, 1) Then
            'ErrorMessage = ErrorString
            GoTo NotFound
        End If
    Next

    If Not Ignore Then
        IParseObject_FindString.Add CStr(LiteralString)
    End If
    
    Exit Function
    
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function

