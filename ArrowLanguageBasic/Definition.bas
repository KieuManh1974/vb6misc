Attribute VB_Name = "Definition"
Option Explicit

Public oParser As ISaffronObject

Public Sub main()
    Dim oFSO As New FileSystemObject
    Dim oContext As clsContext
    Dim sAssembly As String
    
    InitialiseParser
    Set oContext = Program.LoadProgram()
    sAssembly = Intermediate.Compile(oContext)
    Debug.Print sAssembly
    
    oFSO.CreateTextFile(App.Path & "\assembly.txt").Write sAssembly
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
'    SaffronStream.Text = "class int  [ 0 .. 25 ]" & vbCrLf & "class uint [0..255]"
'
'    Dim oTree As New SaffronTree
'
'    If oParser.Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If
End Sub
