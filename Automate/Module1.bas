Attribute VB_Name = "FileScript"
Option Explicit

Public Function LoadScript() As String
    With New FileSystemObject
        LoadScript = .OpenTextFile(App.Path & "\script.txt").ReadAll
    End With
End Function
