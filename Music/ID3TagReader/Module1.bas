Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim oID3 As New clsID3Tag
    Dim oFile As File
    Dim oFSO As New FileSystemObject
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        Set oID3 = New clsID3Tag
        
        oID3.Path = oFile.Path
        Debug.Print oID3.IdealFilename
    Next
End Sub
