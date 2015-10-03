Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    UnzipAll
End Sub

Public Sub UnzipAll()
    Dim oFSO As New FileSystemObject
    Dim oUnZip As New CGUnZipFiles
    Dim lDot As Long
    Dim oFile As File
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        lDot = InStrRev(oFile.Name, ".")
        If UCase$(Mid$(oFile.Name, lDot + 1)) = "ZIP" Then
            oUnZip.Unzip oFile.Name, App.Path
        End If
    Next
End Sub

'Dim oZip As CGZipFiles
'
'Set oZip = New CGZipFiles
'
'oZip.ZipFileName = "\MyZip.Zip"
'oZip.AddFile "c:\mystuff\myfiles\*.*"
'oZip.AddFile "c:\mystuff\mymedia\*.wav"
'
'If oZip.MakeZipFile <> 0 Then
'   MsgBox oZip.GetLastMessage
'End If
'
'Set oZip = Nothing


'The code for Unzipping files is just as straight-forward :

'Set oUnZip = New CGUnZipFiles
'
'
'If oUnZip.UnZip <> 0 Then
'  MsgBox oUnZip.GetLastMessage
'End If
'
'Set oUnZip = Nothing

