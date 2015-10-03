Attribute VB_Name = "Username"
Option Explicit

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As Any, ByVal lpUserName As String, lpnLength As Long) As Long

Public Function GetUserName() As String
    Dim lpUserName As String * 256
    
    WNetGetUser 0&, lpUserName, 256
    GetUserName = UCase$(Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1))
End Function
