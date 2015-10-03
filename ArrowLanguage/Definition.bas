Attribute VB_Name = "Definition"
Option Explicit

Public oParser As ISaffronObject

Public Sub main()
    InitialiseParser
    Program.LoadProgram
End Sub

Public Sub InitialiseParser()
    Dim Definition As String
    Dim oFSO As New FileSystemObject

    Definition = oFSO.OpenTextFile(App.Path & "\arrow.saf").ReadAll

    If Not SaffronObject.CreateRules(Definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = SaffronObject.Rules("statements")

'    Set oParser = SaffronObject.Rules("statements")
'
'    SaffronStream.Text = "for smallint:val{}" & vbCrLf & "var x int"
'
'    Dim oTree As New SaffronTree
'
'    If oParser.Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If
End Sub
