VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parsing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public moReplaceLex As ISaffronObject
Public moSortLex As ISaffronObject
Public moSetLex As ISaffronObject
Public moNewSetLex As ISaffronObject
Public moCamelLex As ISaffronObject

Public Sub InitialiseParser()
    Dim sDef As String
    Dim sPath As String
    
    sPath = App.Path & "\quickpad.saf"
    sDef = Space$(FileLen(sPath))
    Open sPath For Binary As #1
    Get #1, , sDef
    Close #1
    
    If Not CreateRules(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    Set moReplaceLex = Rules("text")
    Set moSortLex = Rules("string")
    Set moSetLex = Rules("set")
    Set moNewSetLex = Rules("element2")
    Set moCamelLex = Rules("camel_text")
    

'    Dim oResult As SaffronTree
'
'    Set oResult = New SaffronTree
'    SaffronStream.Text = "AlphaNumeric23| djjk"
'
'    If moCamelLex.Parse(oResult) Then
'        Stop
'    End If

End Sub

Public Function Decode(ByVal sCode As String) As String
    Dim oTree As SaffronTree
    Dim oSub As SaffronTree
    
    SaffronStream.Text = sCode
    Set oTree = New SaffronTree
    If moReplaceLex.Parse(oTree) Then
        For Each oSub In oTree.SubTree
            Select Case oSub.Index
                Case 1
                    Decode = Decode & "#"
                Case 2 ' #t
                    Decode = Decode & vbTab
                Case 3 ' #n
                    Decode = Decode & vbCrLf
                Case 4 ' #b
                    Decode = Decode & vbCrLf
                Case 5
                    Decode = Decode & Chr(Val(oSub.Text))
                Case 6
                    Decode = Decode & oSub.Text
            End Select
        Next
    Else
        Decode = sCode
    End If
End Function
