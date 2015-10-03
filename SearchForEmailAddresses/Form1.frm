VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moFSO As New FileSystemObject
Private mbSearch As Boolean
Private msSearchFolder As String

Private Sub Form_Initialize()
    msSearchFolder = "c:\"
End Sub


Private Sub cmdSearch_Click()
    If cmdSearch.Caption = "Stop" Then
        mbSearch = False
        cmdSearch.Caption = "Search"
        Exit Sub
    End If
    
    If moFSO.FolderExists(msSearchFolder) Then
        mbSearch = True
        cmdSearch.Caption = "Stop"
        SearchSubFolder moFSO.GetFolder(msSearchFolder)
        cmdSearch.Caption = "Search"
    Else
        txtDirectory.ForeColor = vbRed
    End If

End Sub

Public Sub SearchSubFolder(oFolder As Folder)
    Dim oSubFolder As Folder
    
    On Error GoTo exitsearchsubfolder
    If Not mbSearch Then
        Exit Sub
    End If
    SearchFiles oFolder
    For Each oSubFolder In oFolder.SubFolders
        SearchSubFolder oSubFolder
    Next
exitsearchsubfolder:
End Sub

Public Sub SearchFiles(oFolder As Folder)
    Dim oFile As File
    Dim iDot As Long
    Dim sExtension As String
    
    On Error GoTo ExitSearchFiles
    For Each oFile In oFolder.Files
        With oFile
            If .Size > 0 Then
                iDot = InStrRev(.Name, ".")
                If iDot > 0 Then
                    sExtension = LCase$(Mid$(.Name, iDot + 1))
                    
                    If moExtensions.Exists(sExtension) Then
                        SearchFile oFile
                    End If
                End If
            End If
        End With
        
        DoEvents
        If Not mbSearch Then
            Exit Sub
        End If
    Next
ExitSearchFiles:
End Sub

Public Sub SearchFile(oFile As File)
    Dim oTS As TextStream
    Dim sFile As String
    Dim lDot As Long
    Dim sExtension As String
    Dim lSearchPos As Long
    Dim lSearchPosUnicode As Long
    Dim oListItem As ListItem
    
    On Error GoTo ExitSearchFile
    
    sFile = String$(oFile.Size, "X")
    
    Open oFile.Path For Binary Access Read As #1

    Get #1, , sFile
    Close #1
    
    If Not mbCaseSensitive Then
        sFile = UCase$(sFile)
    End If
    
    lSearchPos = InStr(sFile, msText)
    If lSearchPos = 0 Then
        lSearchPos = InStr(sFile, msUnicodeText)
    End If
    If lSearchPos Then
        lDot = InStrRev(oFile.Path, ".")
        If lDot <> 0 Then
            sExtension = UCase$(Mid$(oFile.Path, lDot + 1))
            If Not moImageExtensions.Exists(sExtension) Then
                imgIcons.ListImages.Add , sExtension, GetIcon(oFile.Path, 16)
                moImageExtensions.Add sExtension, sExtension
            End If
        End If
        
        Debug.Print Mid$(sFile, lSearchPos, 100)
        
        lvwFiles.ListItems.Add , , oFile.Path, , sExtension
        DoEvents
        If Not mbSearch Then
            Exit Sub
        End If
    End If
ExitSearchFile:
    Err.Clear
End Sub


