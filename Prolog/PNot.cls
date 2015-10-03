VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IParseObject

Private ErrorMessage As String
Private NotObject As Object

' // Initialise object as Not
Public Sub Initialise(oObject As IParseObject)
    Set NotObject = oObject
End Sub

Public Function IParseObject_FindString() As Collection
    Dim ReturnCollection As New Collection
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition
    
    Set ReturnCollection = NotObject.IParseObject_FindString
    If ReturnCollection.Count <> 0 Then
        If Not IsNull(ReturnCollection(1)) Then
            GoTo NotFound
        End If
    Else
        GoTo NotFound
    End If
    
    Exit Function
    
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function
