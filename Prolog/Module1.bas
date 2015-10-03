Attribute VB_Name = "Module1"
Option Explicit

Public sParseString As String
Public lStringPosition As Long
Public ParseCollection As New Collection

Public Sub ParseString(sChars As String)
    sParseString = sChars
    lStringPosition = 1
End Sub

Public Function GetChar() As String
    If lStringPosition <= Len(sParseString) Then
        GetChar = Mid$(sParseString, lStringPosition, 1)
        lStringPosition = lStringPosition + 1
    Else
        GetChar = Chr(0)
    End If
End Function
