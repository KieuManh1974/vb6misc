VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParseText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' // Text stream
Private sParseString As String
Private lStringPosition As Long

Private Enum ObjectEnum
    CharacterChoiceObject '
    CharacterRangeObject ' A..Z
    LiteralObject ' ABCD
    JoinObject ' ABCD,DEFG,HIJK
    ChoiceObject 'ABCD|DEFG|HIJK
    RepeatObject '
    RepeatUntilObject '
    NotObject
End Enum

Public Property Let ParseString(sChars As String)
    sParseString = sChars
    lStringPosition = 1
End Property

Private Function GetChar() As String
    If lStringPosition <= Len(sParseString) Then
        GetChar = Mid(sParseString, lStringPosition, 1)
        lStringPosition = lStringPosition + 1
    Else
        GetChar = Chr(0)
    End If
End Function

Public Function IParseObject_FindString(ByVal oCommand As Object) As Collection
    Dim iIndex As Long
    Dim ReturnCollection As Collection
    Dim vMember As Variant
    Dim iTimes As Long
    Dim vReturnMember As Variant
    Dim ThisChar As String * 1
    Dim ReturnString As Variant
    Dim bTerminatorFound As Boolean
    Dim iTempPosition As Long
    Dim AString As String
    Dim bTerminate As Boolean
    Dim ThisPosition As Long
    Dim vElement As Variant
    Dim ErrorMessage As String

    Set IParseObject_FindString = New Collection
    ThisPosition = lStringPosition

    With oCommand
        Select Case .ObjectType

            ' Not operator
            Case NotObject
                Set ReturnCollection = IParseObject_FindString(.ParseCollection)
                If ReturnCollection.Count <> 0 Then
                    If Not IsNull(ReturnCollection(1)) Then
                        GoTo NotFound
                    End If
                Else
                    GoTo NotFound
                End If


            ' Char choice
            Case CharacterChoiceObject
                ThisChar = GetChar
                If InStr(.CharChoice, ThisChar) = 0 Then
                    GoTo NotFound
                End If

                If Not .Ignore Then
                    IParseObject_FindString.Add CStr(ThisChar)
                End If

            ' Literal
            Case LiteralObject
                For iIndex = 1 To Len(.LiteralString)
                    If GetChar <> Mid(.LiteralString, iIndex, 1) Then
                        ErrorMessage = .ErrorString
                        GoTo NotFound
                    End If
                Next

                If Not .Ignore Then
                    IParseObject_FindString.Add CStr(.LiteralString)
                End If

            ' Range
            Case CharacterRangeObject
                ThisChar = GetChar
                If ThisChar < .StartRange Or ThisChar > .EndRange Then
                    ErrorMessage = .ErrorString
                    GoTo NotFound
                End If

                If Not .Ignore Then
                    IParseObject_FindString.Add CStr(ThisChar)
                End If

            ' Repeat
            Case RepeatObject
                If .Complex Then
                    While Not bTerminate And iTimes < .MaximumRepeat
                        For Each vElement In .ParseCollection
                            Set ReturnCollection = IParseObject_FindString(vElement)
                            If ReturnCollection.Count <> 0 Then
                                If Not IsNull(ReturnCollection(1)) Then
                                    If Not .Ignore Then
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

                    If iTimes < .MinimumRepeat Or iTimes > .MaximumRepeat Then
                        ErrorMessage = .ErrorString
                        GoTo NotFound
                    End If
                Else
                    While Not bTerminate And iTimes < .MaximumRepeat
                        For Each vElement In .ParseCollection
                            Set ReturnCollection = IParseObject_FindString(vElement)
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

                    If iTimes < .MinimumRepeat Or iTimes > .MaximumRepeat Then
                        ErrorMessage = .ErrorString
                        GoTo NotFound
                    End If
                    If Not .Ignore Then
                        IParseObject_FindString.Add AString
                    End If
                End If


            ' Repeat Until
            Case RepeatUntilObject
                If .Complex Then
                    While Not IsNull(ReturnString) And Not bTerminatorFound
                        iTempPosition = lStringPosition
                        Set ReturnCollection = IParseObject_FindString(.TerminatorObject)
                        If ReturnCollection.Count = 0 Then
                            lStringPosition = iTempPosition
                            bTerminatorFound = True
                        ElseIf Not IsNull(ReturnCollection(1)) Then
                            lStringPosition = iTempPosition
                            bTerminatorFound = True
                        Else
                            lStringPosition = iTempPosition
                            For Each vElement In .ParseCollection
                                Set ReturnCollection = IParseObject_FindString(vElement)
                                If ReturnCollection.Count <> 0 Then
                                    If Not IsNull(ReturnCollection(1)) Then
                                        IParseObject_FindString.Add ReturnCollection
                                    Else
                                        ErrorMessage = .ErrorString
                                        GoTo NotFound
                                    End If
                                End If
                            Next
                        End If
                    Wend

                Else
                    While Not IsNull(ReturnString) And Not bTerminatorFound
                        iTempPosition = lStringPosition
                        Set ReturnCollection = IParseObject_FindString(.TerminatorObject)
                        If ReturnCollection.Count = 0 Then
                            lStringPosition = iTempPosition
                            bTerminatorFound = True
                        ElseIf Not IsNull(ReturnCollection(1)) Then
                            lStringPosition = iTempPosition
                            bTerminatorFound = True
                        Else
                            lStringPosition = iTempPosition
                            For Each vElement In .ParseCollection
                                Set ReturnCollection = IParseObject_FindString(vElement)
                                If ReturnCollection.Count <> 0 Then
                                    ReturnString = ReturnCollection(1)
                                    If Not IsNull(ReturnString) Then
                                        AString = AString & ReturnString
                                    Else
                                        ErrorMessage = .ErrorString
                                        GoTo NotFound
                                    End If
                                End If
                            Next
                        End If
                    Wend
                    IParseObject_FindString.Add AString
                End If

            ' Choice
            Case ChoiceObject
                If .Complex Then
                    iIndex = 1
                    For Each vElement In .ParseCollection
                        Set ReturnCollection = IParseObject_FindString(vElement)
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
                            If Not .CanIgnore Then
                                ErrorMessage = .ErrorString
                                GoTo NotFound
                            Else
                            End If
                        End If
                    End If
                Else
                    ReturnString = ""
                    iIndex = 1
                    For Each vElement In .ParseCollection
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
                            If Not .CanIgnore Then
                                ErrorMessage = .ErrorString
                                GoTo NotFound
                            Else
                            End If
                        End If
                    End If

                End If

            ' Join
            Case JoinObject
                If .Complex Then
                    For Each vElement In .ParseCollection
                        Set ReturnCollection = IParseObject_FindString(vElement)
                        If ReturnCollection.Count <> 0 Then
                            If IsNull(ReturnCollection(1)) Then
                                GoTo NotFound
                            End If
                            IParseObject_FindString.Add ReturnCollection
                        End If
                    Next
                Else
                    AString = ""
                    For Each vElement In .ParseCollection
                        Set ReturnCollection = IParseObject_FindString(vElement)
                        If ReturnCollection.Count <> 0 Then
                            ReturnString = IParseObject_FindString(vElement)(1)
                            If IsNull(ReturnCollection(1)) Then
                                ErrorMessage = .ErrorString
                                GoTo NotFound
                            End If
                            AString = AString & ReturnCollection(1)
                        End If
                    Next
                    IParseObject_FindString.Add AString
                End If


        End Select
    End With

    Exit Function
NotFound:
    lStringPosition = ThisPosition
    Set IParseObject_FindString = New Collection
    IParseObject_FindString.Add Null
    IParseObject_FindString.Add CStr(ErrorMessage)
    IParseObject_FindString.Add CStr(ThisPosition)
End Function

