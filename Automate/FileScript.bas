Attribute VB_Name = "FileScript"
Option Explicit

Public Function LoadScript(ByVal sFile As String) As String
    With New FileSystemObject
        LoadScript = .OpenTextFile(App.Path & "\" & sFile).ReadAll
    End With
End Function
