VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Search 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text In File"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3615
      Left            =   105
      TabIndex        =   5
      Top             =   1560
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   6376
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imgIcons"
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "test"
         Object.Width           =   21458
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtDirectory 
      Height          =   285
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1200
      Width           =   12135
   End
   Begin VB.TextBox txtSearchText 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   12135
   End
   Begin VB.Label Label2 
      Caption         =   "Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Search Text"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbSearch As Boolean
Private msText As String
Private msUnicodeText As String
Private moExtensions As New Dictionary

Private Sub cmdSearch_Click()
    Dim oFSO As New FileSystemObject
    
    If cmdSearch.Caption = "Stop" Then
        mbSearch = False
        cmdSearch.Caption = "Search"
        Exit Sub
    End If
    
    lvwFiles.ListItems.Clear
    
    If oFSO.FolderExists(txtDirectory.Text) Then
        mbSearch = True
        cmdSearch.Caption = "Stop"
        msText = UCase$(txtSearchText.Text)
        msUnicodeText = Unicode(msText)
        SearchSubFolder oFSO.GetFolder(txtDirectory.Text)
        cmdSearch.Caption = "Search"
    Else
        txtDirectory.ForeColor = vbRed
    End If
End Sub

Private Function Unicode(sText As String) As String
    Dim lIndex As Long
    
    For lIndex = 1 To Len(sText)
        Unicode = Unicode & Chr(0) & Mid$(sText, lIndex, 1)
    Next
End Function

Public Sub SearchSubFolder(oFolder As Folder)
    Dim oSubFolder As Folder
    
    If Not mbSearch Then
        Exit Sub
    End If
    SearchFiles oFolder
    For Each oSubFolder In oFolder.SubFolders
        SearchSubFolder oSubFolder
    Next
End Sub

Public Sub SearchFiles(oFolder As Folder)
    Dim oFile As File
    Dim iDot As Long
    Dim sExtension As String
    
    For Each oFile In oFolder.Files
        With oFile
            If .Size > 0 Then
                iDot = InStr(.Name, ".")
                If iDot > 0 Then
                    sExtension = LCase$(Mid$(.Name, iDot + 1))
                    
                    Select Case sExtension
                        Case "php", "frm", "bas", "cls", "htm", "html", "css"
                            SearchFile oFile
                    End Select
                End If
            End If
        End With
        
        DoEvents
        If Not mbSearch Then
            Exit Sub
        End If
    Next
End Sub

Public Sub SearchFile(oFile As File)
    Dim oTS As TextStream
    Dim sFile As String
    Dim lDot As Long
    Dim sExtension As String
    
    Dim oListItem As ListItem
    
    On Error GoTo ExitSearchFile
    
    sFile = String$(oFile.Size, "X")
    
    'StartCounter
    
    Open oFile.Path For Binary As #1

    Get #1, , sFile
    Close #1
    'Debug.Print GetCounter
    
    sFile = UCase$(sFile)
    
    If InStr(sFile, msText) > 0 Or InStr(sFile, msUnicodeText) Then
        lDot = InStrRev(oFile.Path, ".")
        If lDot <> 0 Then
            sExtension = UCase$(Mid$(oFile.Path, lDot + 1))
            If Not moExtensions.Exists(sExtension) Then
                imgIcons.ListImages.Add , sExtension, GetIcon(oFile.Path, 16)
                moExtensions.Add sExtension, sExtension
            End If
        End If
        
        lvwFiles.ListItems.Add , , oFile.Path, , sExtension
        DoEvents
        If Not mbSearch Then
            Exit Sub
        End If
    End If
ExitSearchFile:
    Err.Clear
End Sub

Private Sub Form_Load()
    lvwFiles.ListItems.Clear
    lvwFiles.ListItems.Add , , "x"
    lvwFiles.ListItems.Clear
End Sub

Private Sub lvwFiles_DblClick()
    Dim sFileName As String
    
    sFileName = lvwFiles.SelectedItem.Text
        
    ShellExecute Me.hDC, "Open", sFileName, "", "", SW_SHOWNORMAL
End Sub

Private Sub txtDirectory_Change()
    txtDirectory.ForeColor = vbBlack
End Sub

Private Sub txtDirectory_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oFSO As New FileSystemObject
    Dim sFolder As String
    
    If Data.GetFormat(vbCFFiles) Then
        sFolder = Data.Files(1)
        If oFSO.FileExists(sFolder) Then
            sFolder = oFSO.GetFile(sFolder).ParentFolder.Path
        End If
            txtDirectory.Text = sFolder
    End If
    
End Sub
