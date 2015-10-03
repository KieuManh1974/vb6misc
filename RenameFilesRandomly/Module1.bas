Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim oFSO As New FileSystemObject
    Dim oFolder As Folder
    Dim oFile As File
    Dim sExt As String
    Dim lDot As Long
    
    Set oFolder = oFSO.GetFolder(App.Path)
    
    For Each oFile In oFolder.Files
        lDot = InStrRev(oFile.Name, ".")
        
        If lDot > 0 Then
            sExt = LCase$(Mid$(oFile.Name, lDot + 1))
            Select Case sExt
                Case "mpg", "avi", "mpeg", "asf", "wmv"
                    oFile.Name = RandomName & "." & sExt
            End Select
        End If
        
    Next
End Sub

Private Function RandomName() As String
    Dim lIndex As Long
    
    For lIndex = 1 To 8
        RandomName = RandomName & Chr$(Int(Rnd * 26) + 65)
    Next
End Function
