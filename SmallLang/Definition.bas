Attribute VB_Name = "Definition"
Option Explicit

Public oParser As ISaffronObject

Public Sub main()
    InitialiseParser
    CompileProgram
End Sub

Public Sub InitialiseParser()
    Dim Definition As String
    Dim oFSO As New FileSystemObject

    Definition = oFSO.OpenTextFile(App.Path & "\smalllang.saf").ReadAll

    If Not SaffronObject.CreateRules(Definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = SaffronObject.Rules("program")
    
    SaffronStream.Text = "a[13]"

    Dim oTree As New SaffronTree

    If oParser.Parse(oTree) Then
        Stop
    Else
        Stop
    End If
End Sub

