VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectSync 
   Caption         =   "ProjectSync"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSync 
      Caption         =   "Sync"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   5400
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwFolders 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   12612
      EndProperty
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   12612
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Folders"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Projects"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmProjectSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moConAccess As New Connection
Private moFSO As New FileSystemObject

Private Sub cmdSync_Click()
    Dim lIndex As Long
    Dim vFolders As Variant
    
    vFolders = Array()
    For lIndex = 1 To lvwFolders.ListItems.Count
        If moFSO.FolderExists(lvwFolders.ListItems(lIndex).Text) Then
            ReDim Preserve vFolders(UBound(vFolders) + 1)
            vFolders(UBound(vFolders)) = lvwFolders.ListItems(lIndex).Text
        End If
    Next
    SyncFiles vFolders
    MsgBox "Files have been synced", vbOKOnly
End Sub

Private Sub Form_Load()
    moConAccess.Open "Provider=Microsoft.Access.OLEDB.10.0;Persist Security Info=False;Data Source=" & App.Path & "/ProjectSync.mdb;User ID=Admin;Data Provider=Microsoft.Jet.OLEDB.4.0"
    LoadProjects

    lvwFolders.ListItems.Add , , "New Folder"
End Sub

Private Sub lvwFolders_AfterLabelEdit(Cancel As Integer, sNewString As String)
    Dim sCurrentName As String
    
    If lvwFolders.SelectedItem.Index = lvwFolders.ListItems.Count Then
        AddFolder sNewString, lvwProjects.SelectedItem.Index
        lvwFolders.ListItems.Add , , "New Folder"
    Else
        RenameFolder sNewString, lvwFolders.SelectedItem.Text, lvwProjects.SelectedItem.Index
    End If
End Sub

Private Sub lvwProjects_AfterLabelEdit(Cancel As Integer, sNewString As String)
    Dim sCurrentName As String
    
    If lvwProjects.SelectedItem.Index = lvwProjects.ListItems.Count Then
        AddProject sNewString
        lvwProjects.ListItems.Add , , "New Project"
    Else
        RenameProject sNewString, lvwProjects.SelectedItem.Index
    End If
End Sub

Private Sub LoadProjects()
    Dim sSQL As String
    Dim oProject As New Recordset
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     Name"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     Projects"
    sSQL = sSQL & " ORDER BY"
    sSQL = sSQL & "     ProjectIndex"

    oProject.Open sSQL, moConAccess, adOpenForwardOnly, , adCmdText
    
    While Not oProject.EOF
        lvwProjects.ListItems.Add , , oProject!Name
        oProject.MoveNext
    Wend
    
    lvwProjects.ListItems.Add , , "New Project"
End Sub

Private Sub LoadFolders(ByVal lProjectIndex As Long)
    Dim sSQL As String
    Dim oFolder As New Recordset
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     Name"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     Folders"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     ProjectIndex = " & lProjectIndex

    oFolder.Open sSQL, moConAccess, adOpenForwardOnly, , adCmdText
    
    lvwFolders.ListItems.Clear
    While Not oFolder.EOF
        lvwFolders.ListItems.Add , , oFolder!Name
        oFolder.MoveNext
    Wend
    
    lvwFolders.ListItems.Add , , "New Folder"
End Sub

Private Sub RenameFolder(sNewName As String, sOldName As String, lProjectIndex)
    Dim oUpdate As New Recordset
    Dim sSQL As String
    
    sSQL = ""
    sSQL = sSQL & " UPDATE"
    sSQL = sSQL & "     Folders"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "     Name = '" & sNewName & "'"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     Name = '" & sOldName & "'"
    sSQL = sSQL & "     AND ProjectIndex = " & lProjectIndex
    
    oUpdate.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
End Sub

Private Sub RenameProject(sNewName As String, lProjectIndex As Long)
    Dim oUpdate As New Recordset
    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL & " UPDATE"
    sSQL = sSQL & "     Projects"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "     Name = '" & sNewName & "'"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     ProjectIndex = " & lProjectIndex

    oUpdate.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
End Sub

Private Sub AddProject(sNewProject As String)
    Dim lMaxIndex As Long
    Dim oMaxIndex As New Recordset
    Dim oInsert As New Recordset
    Dim sSQL As String
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     MAX(ProjectIndex) AS MaxIndex"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     Projects"
        
    oMaxIndex.Open sSQL, moConAccess, adOpenForwardOnly, , adCmdText
    
    If Not IsNull(oMaxIndex!MaxIndex) Then
        lMaxIndex = oMaxIndex!MaxIndex + 1
    Else
        lMaxIndex = 1
    End If
    
    sSQL = ""
    sSQL = sSQL & " INSERT INTO"
    sSQL = sSQL & "     Projects"
    sSQL = sSQL & "     (Name,"
    sSQL = sSQL & "     ProjectIndex)"
    sSQL = sSQL & " VALUES"
    sSQL = sSQL & "     ('" & sNewProject & "',"
    sSQL = sSQL & "     " & lMaxIndex & ")"
    
    oInsert.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
End Sub

Private Function AddFolder(sNewFolder As String, lProjectIndex As Long)
    Dim oInsert As New Recordset
    Dim sSQL As String
        
    sSQL = ""
    sSQL = sSQL & " INSERT INTO"
    sSQL = sSQL & "     Folders"
    sSQL = sSQL & "     (Name,"
    sSQL = sSQL & "     ProjectIndex)"
    sSQL = sSQL & " VALUES"
    sSQL = sSQL & "     ('" & sNewFolder & "',"
    sSQL = sSQL & "     " & lProjectIndex & ")"
    
    oInsert.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
End Function

Private Sub lvwProjects_Click()
    If lvwProjects.SelectedItem.Index < lvwProjects.ListItems.Count Then
        LoadFolders lvwProjects.SelectedItem.Index
    Else
        lvwFolders.ListItems.Clear
    End If
End Sub

Private Sub SyncFiles(ByVal vFolders As Variant)
    Dim vFiles As Variant
    Dim vFolder1 As Variant
    Dim vFolder2 As Variant
    Dim vFile1 As Variant
    Dim vFile2 As Variant
    
    Dim oFile As File
    Dim vSubFiles As Variant

    Dim dRecentDate As Date

    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim bOK As Boolean
    Dim vLatestFiles As Variant
    
    Const FOLDER_NAME = 0
    Const FOLDER_FILES = 1
    Const FILE_NAME = 0
    Const FILE_DATE = 1
    Const FILE_FOLDER = 2
    
    vLatestFiles = Array()
    vFiles = Array()
    
    For Each vFolder1 In vFolders
        ReDim Preserve vFiles(UBound(vFiles) + 1)
        vFiles(UBound(vFiles)) = Array(CStr(vFolder1), Array())
        vSubFiles = Array()
        For Each oFile In moFSO.GetFolder(vFolder1).Files
            ReDim Preserve vSubFiles(UBound(vSubFiles) + 1)
            vSubFiles(UBound(vSubFiles)) = Array(oFile.Name, oFile.DateLastModified, CStr(vFolder1))
        Next
        vFiles(UBound(vFiles))(1) = vSubFiles
    Next
    
    ' Find latest files to copy
    For lIndex1 = 0 To UBound(vFiles) - 1
        For lIndex2 = lIndex1 + 1 To UBound(vFiles)
            If lIndex1 <> lIndex2 Then
                vFolder1 = vFiles(lIndex1)
                vFolder2 = vFiles(lIndex2)
                For Each vFile1 In vFolder1(FOLDER_FILES)
                    For Each vFile2 In vFolder2(FOLDER_FILES)
                        If vFile1(FILE_NAME) = vFile2(FILE_NAME) Then
                            AddRecent vLatestFiles, vFile1
                            AddRecent vLatestFiles, vFile2
                        End If
                    Next
                Next
            End If
        Next
    Next
    
    For lIndex1 = 0 To UBound(vFiles)
        vFolder1 = vFiles(lIndex1)
        For Each vFile1 In vFolder1(FOLDER_FILES)
            For Each vFile2 In vLatestFiles
                If vFile1(FILE_NAME) = vFile2(FILE_NAME) Then
                    If vFile1(FILE_FOLDER) <> vFile2(FILE_FOLDER) Then
                        On Error Resume Next
                        moFSO.CopyFile vFile2(FILE_FOLDER) & "/" & vFile2(FILE_NAME), vFile1(FILE_FOLDER) & "/" & vFile1(FILE_NAME)
                        If Err.Description <> "" Then
                            MsgBox Err.Description & vbCrLf & vFile2(FILE_FOLDER) & "/" & vFile2(FILE_NAME) & vbCrLf & vFile1(FILE_FOLDER) & "/" & vFile1(FILE_NAME), vbOKOnly, "Error"
                            Err.Clear
                        End If
                    End If
                End If
            Next
        Next
    Next
    
End Sub

Private Function AddRecent(vMostRecentFiles As Variant, vFile As Variant)
    Dim lIndex As Long
    Const FILE_NAME = 0
    Const FILE_DATE = 1
    Const FILE_FOLDER = 2
    
    For lIndex = 0 To UBound(vMostRecentFiles)
        If vMostRecentFiles(lIndex)(FILE_NAME) = vFile(FILE_NAME) Then
            If vFile(FILE_DATE) > vMostRecentFiles(lIndex)(FILE_DATE) Then
                vMostRecentFiles(lIndex) = vFile
                Exit Function
            Else
                Exit Function
            End If
        End If
    Next
    
    ReDim Preserve vMostRecentFiles(UBound(vMostRecentFiles) + 1)
    vMostRecentFiles(UBound(vMostRecentFiles)) = vFile
End Function

Private Sub lvwProjects_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sSQL As String
    Dim oDelete As New Recordset
    Dim lProjectIndex As Long
    
    If KeyCode = 46 Then
        lProjectIndex = lvwProjects.SelectedItem.Index
        
        If MsgBox("Ok to delete?", vbYesNo) = vbYes Then
            sSQL = ""
            sSQL = sSQL & " DELETE"
            sSQL = sSQL & " FROM"
            sSQL = sSQL & "     Projects"
            sSQL = sSQL & " WHERE"
            sSQL = sSQL & "     ProjectIndex = " & lProjectIndex
            oDelete.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
        
            sSQL = ""
            sSQL = sSQL & " DELETE"
            sSQL = sSQL & " FROM"
            sSQL = sSQL & "     Folders"
            sSQL = sSQL & " WHERE"
            sSQL = sSQL & "     ProjectIndex = " & lProjectIndex
            oDelete.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
            lvwProjects.ListItems.Remove lProjectIndex
            lvwFolders.ListItems.Clear
        End If
    End If
End Sub

Private Sub lvwFolders_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sSQL As String
    Dim oDelete As New Recordset
    Dim lProjectIndex As Long
    
    If KeyCode = 46 Then
        If MsgBox("Ok to delete?", vbYesNo) = vbYes Then
            sSQL = ""
            sSQL = sSQL & " DELETE"
            sSQL = sSQL & " FROM"
            sSQL = sSQL & "     Folders"
            sSQL = sSQL & " WHERE"
            sSQL = sSQL & "     Name = '" & lvwFolders.SelectedItem.Text & "'"
            oDelete.Open sSQL, moConAccess, , adLockOptimistic, adCmdText
            lvwFolders.ListItems.Remove lvwProjects.SelectedItem.Index
        End If
    End If
End Sub
