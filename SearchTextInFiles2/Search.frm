VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Search 
   Caption         =   "Search Text In File"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExtensions 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   5280
      Width           =   9495
   End
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
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   6376
      View            =   3
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
         Object.Width           =   21361
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
   Begin VB.Label Label3 
      Caption         =   "Extensions :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   975
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

Private mlMouseShift As Long
Private moFSO As New FileSystemObject
Private msExtensions As String
Private moParser As ISaffronObject

Private moExtensions As New Dictionary
Private moImageExtensions As New Dictionary

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Sub cmdSearch_Click()
    If cmdSearch.Caption = "Stop" Then
        mbSearch = False
        cmdSearch.Caption = "Search"
        Exit Sub
    End If
    
    lvwFiles.ListItems.Clear
    
    If moFSO.FolderExists(txtDirectory.Text) Then
        mbSearch = True
        cmdSearch.Caption = "Stop"
        msText = UCase$(txtSearchText.Text)
        msUnicodeText = StrConv(msText, vbUnicode)
        SearchSubFolder moFSO.GetFolder(txtDirectory.Text)
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
    
    Open oFile.Path For Binary As #1

    Get #1, , sFile
    Close #1
    
    sFile = UCase$(sFile)
    
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

Private Sub Form_Initialize()
    Dim vExtensions As Variant
    Dim vExtension As Variant
    Dim sDef As String
    
    sDef = "extension list in 0 to 9 a to z _ | | | "
    sDef = sDef & "noex omit list in 000 to 255 not 0 to 9 not a to z not _ | | | "
    sDef = sDef & "extensions list or extension noex | | | "
    
    If Not CreateRules(sDef) Then
        MsgBox "Bad def"
        End
    End If
    Set moParser = SaffronCompiler.Rules("extensions")
    
    msExtensions = GetSetting("SearchTextInFiles", "Preferences", "Extensions", "php,frm,bas,cls,htm,html,css,xml,xsl,inc")
    vExtensions = Split(msExtensions, ",")
    
    txtExtensions.Text = Join(vExtensions, " ")
    Set moExtensions = New Dictionary
    
    For Each vExtension In vExtensions
        moExtensions.Add CStr(vExtension), CStr(vExtension)
    Next
End Sub

Private Sub Form_Load()
    lvwFiles.ListItems.Clear
    lvwFiles.ListItems.Add , , "x"
    lvwFiles.ListItems.Clear
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub lvwFiles_DblClick()
    Dim sFileName As String
    Dim lReturn As Long
    
    sFileName = lvwFiles.SelectedItem.Text
    If mlMouseShift = 0 Then
         lReturn = ShellExecute(Me.hDC, "Open", sFileName, "", "", SW_SHOWNORMAL)
         If lReturn < 32 Then
            MsgBox "Error " & lReturn
        End If
    Else
        If moFSO.FileExists(sFileName) Then
            sFileName = moFSO.GetFile(sFileName).ParentFolder.Path
        End If
        
        CreateProcessX "explorer.exe /e," & sFileName
    End If
End Sub

Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlMouseShift = Shift
End Sub

Private Sub txtDirectory_Change()
    txtDirectory.ForeColor = vbBlack
End Sub

Private Sub txtDirectory_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sFolder As String
    
    If Data.GetFormat(vbCFFiles) Then
        sFolder = Data.Files(1)
        If moFSO.FileExists(sFolder) Then
            sFolder = moFSO.GetFile(sFolder).ParentFolder.Path
        End If
            txtDirectory.Text = sFolder
    End If
    
End Sub

Private Sub txtExtensions_LostFocus()
    Dim oRuleTree As SaffronTree
    Dim oSubTree As SaffronTree
    Dim vExtensions As Variant
    Dim vExtension As Variant
    
    vExtensions = Array()
    
    Set oRuleTree = New SaffronTree
    
    SaffronStream.Text = txtExtensions.Text
    If moParser.Parse(oRuleTree) Then
        For Each oSubTree In oRuleTree.SubTree
            If oSubTree.index = 1 Then
                ReDim Preserve vExtensions(UBound(vExtensions) + 1)
                vExtensions(UBound(vExtensions)) = LCase$(oSubTree.Text)
            End If
        Next
    End If
    SaveSetting "SearchTextInFiles", "Preferences", "Extensions", Join(vExtensions, ",")
    txtExtensions.Text = Join(vExtensions, " ")
    
    Set moExtensions = New Dictionary
    
    For Each vExtension In vExtensions
        moExtensions.Add CStr(vExtension), CStr(vExtension)
    Next
End Sub
