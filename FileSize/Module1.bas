Attribute VB_Name = "Module1"
Option Explicit

Dim MyInfo As New sizes

Sub Main()
    Dim ofso As New FileSystemObject
    Dim q As Long
    Recurse ofso.GetFolder("c:\")
    
    Dim sa As Long
    
    sa = MyInfo.MaxArray
    
    Open "c:\sizes.txt" For Output As #1
    
    With MyInfo
        For q = 0 To sa
            Print #1, .GetItemSize(q) & "," & .GetItemOccurance(q)
        Next
    End With
    Close #1
End Sub

Private Sub Recurse(oFolder As Folder)
    Dim oFile As File
    Dim oSubFolder As Folder
    
    For Each oFile In oFolder.Files
        MyInfo.Include oFile.Size
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        Recurse oSubFolder
    Next
End Sub
