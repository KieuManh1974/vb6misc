VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRepeat"
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

' // For repeat collection
Private MinimumRepeat As Long
Private MaximumRepeat As Long


Private ErrorMessage As String

Private ParseCollection As New Collection

' // Initialise object as Repeat
Public Sub Initialise(bIgnore As Boolean, bComplex As Boolean, MinR As Long, MaxR As Long, ParamArray vElements() As Variant)
    MinimumRepeat = MinR
    MaximumRepeat = MaxR
    
    Dim iIndex As Long
    For iIndex = LBound(vElements) To UBound(vElements)
        ParseCollection.Add vElements(iIndex)
    Next

    Complex = bComplex
    Ignore = bIgnore
End Sub

Public Function IParseObject_FindString() As Collection
    Dim bTerminate As Boolean
    Dim iTimes As Integer
    Dim vElement As IParseObject
    Dim ReturnCollection As Collection
    Dim vReturnMember As Collection
    Dim AString As String
    Dim ThisPosition As Long
    
    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition
    
    If Complex Then
        While Not bTerminate And iTimes < MaximumRepeat
            For Each vElement In ParseCollection
                Set ReturnCollection = IParseObject_FindString(vElement)
                If ReturnCollection.Count <> 0 Then
                    If Not IsNull(ReturnCollection(1)) Then
                        If Not Ignore Then
                            For Each vReturnMember In ReturnCollection
                                IParseObject_FindString.Add vReturnMember
                            Next
                        End If
                        iTimes = iTimes + 1
                    Else
                        bTerminate = True
                    End If
                End If
            Next
        Wend

        If iTimes < MinimumRepeat Or iTimes > MaximumRepeat Then
            'ErrorMessage = .ErrorString
            GoTo NotFound
        End If
    Else
        While Not bTerminate And iTimes < MaximumRepeat
            For Each vElement In ParseCollection
                Set ReturnCollection = vElement.FindString
                If ReturnCollection.Count <> 0 Then
                    If Not IsNull(ReturnCollection(1)) Then
                        AString = AString & ReturnCollection(1)
                        iTimes = iTimes + 1
                    Else
                        bTerminate = True
                    End If
                End If
            Next
        Wend

        If iTimes < MinimumRepeat Or iTimes > MaximumRepeat Then
            'ErrorMessage = .ErrorString
            GoTo NotFound
        End If
        If Not Ignore Then
            IParseObject_FindString.Add AString
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

