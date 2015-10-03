VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Search 
   Caption         =   "Search Text In File"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cboExtensions 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   5880
      Width           =   7455
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   4320
      Top             =   0
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
      Top             =   1680
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "imgIcons"
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   21361
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtDirectory 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1320
      Width           =   12135
   End
   Begin VB.TextBox txtSearchText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   12135
   End
   Begin VB.Label lblCurrentPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   12135
   End
   Begin VB.Label lblExtensions 
      Caption         =   "Extensions :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Search Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetFileInformationByHandle Lib "kernel32.dll" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Const OFS_MAXPATHNAME = 128
Private Const OF_READ = &H0
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type

Private mbSearch As Boolean
Private msText As String
Private msUnicodeText As String

Private mlMouseShift As Long
Private moFSO As New FileSystemObject
Private msExtensions As String
Private moParser As ISaffronObject
Private moParserSearch As ISaffronObject

Private moExtensions As New Dictionary
Private mvExtensionList As Variant
Private mlSelectedExtensions As Long
Private moImageExtensions As New Dictionary

Private mnWidth As Single
Private mnComboWidth As Single
Private mnSearchLeft As Single
Private mnAddLeft As Single
Private mnRemoveLeft As Single
Private mnHeight As Single
Private mnRowTop As Single
Private mnRowTop2 As Single

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

Private mbCaseSensitive As Boolean

Private Type FileMatch
    FileId As Long
    Pairs(255, 31) As Byte
End Type

Private mfmFileMatches() As FileMatch
Private mlCountFileMatches As Long

Private myValues(7) As Byte

Private Sub chkCaseSensitive_Click()
    mbCaseSensitive = chkCaseSensitive.Value = vbChecked
End Sub

Private Sub cmdAdd_Click()
    cboExtensions.AddItem cboExtensions.Text
End Sub

Private Sub cmdCopy_Click()
    Dim oList As ListItem
    Dim sText As String
    
    For Each oList In lvwFiles.ListItems
        sText = sText & vbCrLf & oList.Text
    Next
    
    If sText <> "" Then
        Clipboard.Clear
        Clipboard.SetText Mid$(sText, 3), vbCFText
    End If
End Sub

Private Sub cmdRemove_Click()
    If cboExtensions.ListIndex <> -1 Then
        cboExtensions.RemoveItem cboExtensions.ListIndex
    End If
End Sub

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
        
        If Not mbCaseSensitive Then
            msText = UCase$(Decode(txtSearchText.Text))
        Else
            msText = Decode(txtSearchText.Text)
        End If
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
    
        lblCurrentPath.Caption = oFile.Path
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
    
    Dim fiInfo As BY_HANDLE_FILE_INFORMATION
    Dim hFile As Long
    Dim lpReOpenBuff As OFSTRUCT
    Dim Ret As Long
    
    Dim yFile() As Byte
    Dim lFileMatchIndex As Long
    
    On Error GoTo ExitSearchFile
    
    sFile = String$(oFile.Size, "X")
    
    
    hFile = OpenFile(oFile.Path, lpReOpenBuff, OF_READ)
    Ret = GetFileInformationByHandle(hFile, fiInfo)
    
    lFileMatchIndex = GetFileMatch(fiInfo.nFileIndexLow)
    If lFileMatchIndex <> -1 Then
        If HasAllMatches(mfmFileMatches(lFileMatchIndex), msUnicodeText) Then
            
        End If
    Else
        
    End If

    Open oFile.Path For Binary Access Read As #1
    ReDim yFile(oFile.Size - 1)
    Get #1, , yFile
    BuildPairs fiInfo.nFileIndexLow, yFile, oFile.Size

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

Private Sub BuildPairs(ByVal lFileId As Long, yFile() As Byte, lLength As Long)
    Dim lIndex As Long
    Dim yPairs(255, 31) As Byte
    Dim lX As Long
    Dim lY As Long
    Dim lXVal As Byte
    Dim lYVal As Byte
    
    ReDim Preserve mfmFileMatches(mlCountFileMatches)
    mfmFileMatches(mlCountFileMatches).FileId = lFileId
    
    For lIndex = 0 To lLength - 2
        lX = yFile(lIndex) \ 8
        lY = yFile(lIndex + 1) \ 8
        lYVal = myValues(yFile(lIndex + 1) And 7)
        
        mfmFileMatches(mlCountFileMatches).Pairs(lX, lY) = mfmFileMatches(mlCountFileMatches).Pairs(lX, lY) Or lYVal
    Next
    
    mlCountFileMatches = mlCountFileMatches + 1
End Sub

Private Function HasAllMatches(fmFileMatch As FileMatch, sSearch As String) As Boolean
    Dim lIndex  As Long
    
    For lIndex = 1 To Len(sSearch) - 1
        If Not HasMatch(fmFileMatch, Asc(Mid$(sSearch, lIndex, 1)), Asc(Mid$(sSearch, lIndex + 1, 1))) Then
            Exit Function
        End If
    Next
    HasAllMatches = True
End Function

Private Function HasMatch(fmFileMatch As FileMatch, yByte1 As Byte, yByte2 As Byte) As Boolean
    Dim lX As Long
    Dim lY As Long
    Dim lYVal As Byte
    
    lX = yByte1 \ 8
    lY = yByte2 \ 8
    lYVal = myValues(yByte2 And 7)
    
    HasMatch = (fmFileMatch.Pairs(lX, lY) And lYVal) <> 0
End Function

Private Function GetFileMatch(ByVal lFileId As Long) As Long
    Dim lIndex As Long
    For lIndex = 0 To mlCountFileMatches - 1
        If mfmFileMatches(lIndex).FileId = lFileId Then
            GetFileMatch = lIndex
            Exit Function
        End If
    Next
    GetFileMatch = -1
End Function

Private Sub Form_Initialize()
    Dim vExtensions As Variant
    Dim vExtension As Variant
    Dim lIndex As Long
    Dim vSplit As Variant
    Dim sDef As String

    Dim yValue As Long
    
    yValue = 1
    For lIndex = 0 To 7
        myValues(lIndex) = yValue
        yValue = yValue * 2
    Next
    
    sDef = "extension list in 0 to 9 a to z _ | | | "
    sDef = sDef & "noex omit list in 000 to 255 not 0 to 9 not a to z not _ | | | "
    sDef = sDef & "extensions list or extension noex | | | "
    sDef = sDef & "text list or ## #t #n #b and omit # list in 0 to 9 | min 3 max 3 | | skip | | until eos | | "
    
    If Not CreateRules(sDef) Then
        MsgBox "Bad def"
        End
    End If
    Set moParser = SaffronCompiler.Rules("extensions")
    Set moParserSearch = SaffronCompiler.Rules("text")
    
    msExtensions = GetSetting("SearchTextInFiles", "Preferences", "Extensions", "php,frm,bas,cls,htm,html,css,xml,xsl,inc")
    mlSelectedExtensions = GetSetting("SearchTextInFiles", "Preferences", "SelectedExtension", 0)
    
    vSplit = Split(msExtensions, "|")
    ReDim mvExtensionList(UBound(vSplit))
    For lIndex = 0 To UBound(vSplit)
        mvExtensionList(lIndex) = Split(vSplit(lIndex), ",")
    Next
    
    cboExtensions.Text = Join(mvExtensionList(mlSelectedExtensions), " ")
    
    Set moExtensions = New Dictionary
    
    On Error Resume Next
    For Each vExtension In mvExtensionList(mlSelectedExtensions)
        moExtensions.Add CStr(vExtension), CStr(vExtension)
    Next
    
End Sub

Public Function Decode(ByVal sCode As String) As String
    Dim oTree As SaffronTree
    Dim oSub As SaffronTree
    
    SaffronStream.Text = sCode
    Set oTree = New SaffronTree
    If moParserSearch.Parse(oTree) Then
        For Each oSub In oTree.SubTree
            Select Case oSub.index
                Case 1
                    Decode = Decode & "#"
                Case 2 ' #t
                    Decode = Decode & vbTab
                Case 3 ' #n
                    Decode = Decode & vbCrLf
                Case 4 ' #b
                    Decode = Decode & vbCrLf
                Case 5
                    Decode = Decode & Chr(Val(oSub.Text))
                Case 6
                    Decode = Decode & oSub.Text
            End Select
        Next
    Else
        Decode = sCode
    End If
End Function

Private Sub Form_Load()
    lvwFiles.ListItems.Clear
    lvwFiles.ListItems.Add , , "x"
    lvwFiles.ListItems.Clear
    
    mnWidth = Me.Width - txtSearchText.Width
    mnComboWidth = Me.Width - cboExtensions.Width
    mnSearchLeft = Me.Width - cmdSearch.Left
    mnAddLeft = Me.Width - cmdAdd.Left
    mnRemoveLeft = Me.Width - cmdRemove.Left
    mnHeight = Me.Height - lvwFiles.Height
    mnRowTop = Me.Height - cboExtensions.Top
    mnRowTop2 = Me.Height - lblCurrentPath.Top
    
    Me.Width = GetSetting("SearchTextInFiles", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("SearchTextInFiles", "Dimensions", "Height", Me.Height)
    
    Me.Top = GetSetting("SearchTextInFiles", "Dimensions", "Top", Me.Top)
    Me.Left = GetSetting("SearchTextInFiles", "Dimensions", "Left", Me.Left)
    
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    If Me.Left < 0 Then
        Me.Left = 0
    End If
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtSearchText.Width = Me.Width - mnWidth
    txtDirectory.Width = Me.Width - mnWidth
    lvwFiles.Width = Me.Width - mnWidth
    cboExtensions.Width = Me.Width - mnComboWidth
    cmdSearch.Left = Me.Width - mnSearchLeft
    cmdAdd.Left = Me.Width - mnAddLeft
    cmdRemove.Left = Me.Width - mnRemoveLeft
    
    lblCurrentPath.Top = Me.Height - mnRowTop2
    lvwFiles.Height = Me.Height - mnHeight
    cboExtensions.Top = Me.Height - mnRowTop
    lblExtensions.Top = Me.Height - mnRowTop
    cmdSearch.Top = Me.Height - mnRowTop
    cmdAdd.Top = Me.Height - mnRowTop
    cmdRemove.Top = Me.Height - mnRowTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "SearchTextInFiles", "Dimensions", "Width", Me.Width
    SaveSetting "SearchTextInFiles", "Dimensions", "Height", Me.Height
    SaveSetting "SearchTextInFiles", "Dimensions", "Top", Me.Top
    SaveSetting "SearchTextInFiles", "Dimensions", "Left", Me.Left
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

Private Sub lvwFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim vFile As Variant
    Dim lDot As Long
    Dim sExtension As String
    
    lvwFiles.ListItems.Clear
    For Each vFile In Data.Files
        lDot = InStrRev(vFile, ".")
        If lDot <> 0 Then
            sExtension = UCase$(Mid$(vFile, lDot + 1))
        Else
            sExtension = "folder"
        End If
        
        If Not moImageExtensions.Exists(sExtension) Then
            imgIcons.ListImages.Add , sExtension, GetIcon(CStr(vFile), 16)
            moImageExtensions.Add sExtension, sExtension
        End If
        lvwFiles.ListItems.Add , , vFile, , sExtension
    Next
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

Private Sub cboExtensions_LostFocus()
    Dim oRuleTree As SaffronTree
    Dim oSubTree As SaffronTree
    Dim vExtensions As Variant
    Dim vExtension As Variant
    
    vExtensions = Array()
    
    Set oRuleTree = New SaffronTree
    
    SaffronStream.Text = cboExtensions.Text
    If moParser.Parse(oRuleTree) Then
        For Each oSubTree In oRuleTree.SubTree
            If oSubTree.index = 1 Then
                ReDim Preserve vExtensions(UBound(vExtensions) + 1)
                vExtensions(UBound(vExtensions)) = LCase$(oSubTree.Text)
            End If
        Next
    End If
    
    SaveSetting "SearchTextInFiles", "Preferences", "Extensions", Join(vExtensions, ",")
    cboExtensions.Text = Join(vExtensions, " ")
    
    Set moExtensions = New Dictionary
    
    For Each vExtension In vExtensions
        moExtensions.Add CStr(vExtension), CStr(vExtension)
    Next
End Sub
