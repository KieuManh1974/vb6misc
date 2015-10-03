VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moParser As ISaffronObject

Public Function HintText2(sText As String) As String
'    Dim sFormatted As String
'
'    '{\rtf1\ansi\deff0{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}{\cf1 hello}}
'    If sText = "function" Then
'        sFormatted = "function Name() \\ \{\}"
'    Else
'        sFormatted = sText
'    End If
'
'    sFormatted = Replace$(sFormatted, "\", "\\")
'    sFormatted = Replace$(sFormatted, "{", "\{")
'    sFormatted = Replace$(sFormatted, "}", "\}")
'
'    HintText = "{\rtf1\ansi\{\colortbl;\red255\green0\blue0\} " & sFormatted & "}"
'    HintText = "{\rtf1\ansi\deff0{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}{\cf1 hello}}"
End Function

Public Function HintText(sText As String) As String
    Dim oTree As SaffronTree
    
    SaffronStream.Text = sText
    Set oTree = New SaffronTree
    
    If Not moParser.Parse(oTree) Then
        CompleteHTML oTree
    End If
End Function

Private Function TreeToRTF(oTree As SaffronTree) As String
    
End Function
    
Private Sub Class_Initialize()
    LoadRules
End Sub

Private Sub LoadRules()
    Dim sDefinition As String
    
    sDefinition = Space$(FileLen(App.Path & "\syntaxhighlight.saf"))
    Open App.Path & "\syntaxhighlight.saf" For Binary As #1
    Get #1, , sDefinition
    Close #1

    If Not CreateRules(sDefinition) Then
        Debug.Print ErrorString
        End
    End If
    
'    Dim oTree As SaffronTree
'
'    Set oTree = New SaffronTree
'    SaffronStream.Text = "<table>"
'    If Rules("table").Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If

    Set moParser = Rules("html")
End Sub
