VERSION 5.00
Begin VB.Form frmChooser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Code Fragment"
   ClientHeight    =   11865
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11865
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11295
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
   Begin VB.ListBox lstTitles 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11190
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nOriginalWidth As Single
Private moTitleList As New clsNode
Private moDatabase As New Connection

Private Sub Form_Load()

    nOriginalWidth = Me.Width - txtCode.Width
    moDatabase.Open "Provider=Microsoft.Access.OLEDB.10.0;Persist Security Info=False;Data Source=" & App.Path & "\CodeFragment.mdb;User ID=Admin;Data Provider=Microsoft.Jet.OLEDB.4.0"
    LoadFragments
    
    Me.Height = GetSetting("CodeFragment", "Dimensions", "Height", Me.Height)
    Me.Width = GetSetting("CodeFragment", "Dimensions", "Width", Me.Width)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFragments
    SaveSetting "CodeFragment", "Dimensions", "Height", Me.Height
    SaveSetting "CodeFragment", "Dimensions", "Width", Me.Width
End Sub

Private Sub Form_Resize()
    lstTitles.Height = Me.Height
    txtCode.Height = Me.Height
    txtCode.Width = Me.Width - nOriginalWidth
End Sub

Private Sub LoadFragments()
    Dim sSQL As String
    Dim oFragments As New Recordset
    Dim oNode As clsNode
    Dim oInfo As clsInfo
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     FragmentIndex,"
    sSQL = sSQL & "     Title,"
    sSQL = sSQL & "     CodeFragment"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     CodeFragments"
    sSQL = sSQL & " ORDER BY"
    sSQL = sSQL & "     OrderIndex"
    
    oFragments.Open sSQL, moDatabase, adOpenForwardOnly, , adCmdText
    
    Set moTitleList = New clsNode
    lstTitles.Clear
    
    While Not oFragments.EOF
        Set oNode = moTitleList.AddNew(, oFragments!FragmentIndex)
        Set oInfo = New clsInfo
        oInfo.Title = oFragments!Title
        oInfo.Text = oFragments!CodeFragment
        lstTitles.AddItem oFragments!Title
        Set oNode.Value = oInfo
        oFragments.MoveNext
    Wend
End Sub

Private Sub SaveFragments()
    Dim lIndex As Long
    Dim oNode As clsNode
    Dim oFragments As New Recordset
    Dim sSQL As String
    Dim sDeletes As String
    
    For lIndex = 0 To moTitleList.Count - 1
        Set oNode = moTitleList.ItemPhysical(lIndex)
        sDeletes = sDeletes & "," & oNode.LogicalKey
        
        sSQL = ""
        sSQL = sSQL & " SELECT"
        sSQL = sSQL & "     FragmentIndex,"
        sSQL = sSQL & "     Title,"
        sSQL = sSQL & "     CodeFragment,"
        sSQL = sSQL & "     OrderIndex"
        sSQL = sSQL & " FROM"
        sSQL = sSQL & "     CodeFragments"
        sSQL = sSQL & " WHERE"
        sSQL = sSQL & "     FragmentIndex = " & oNode.LogicalKey

        Set oFragments = New Recordset
        oFragments.Open sSQL, moDatabase, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Not oFragments.EOF Then
            oFragments!Title = oNode.Value.Title
            oFragments!CodeFragment = oNode.Value.Text
            oFragments!OrderIndex = lIndex
            oFragments.Update
        Else
            oFragments.AddNew
            oFragments!FragmentIndex = oNode.LogicalKey
            oFragments!Title = oNode.Value.Title
            oFragments!CodeFragment = oNode.Value.Text
            oFragments!OrderIndex = lIndex
            oFragments.Update
        End If
    Next
    sDeletes = Mid$(sDeletes, 2)
    sSQL = ""
    sSQL = sSQL & " DELETE"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     CodeFragments"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     FragmentIndex NOT IN (" & sDeletes & ")"
    Set oFragments = New Recordset
    oFragments.Open sSQL, moDatabase, adOpenDynamic, adLockOptimistic, adCmdText
    
End Sub

Private Sub lstTitles_Click()
    Dim oNode As clsNode
    
    If lstTitles.ListIndex <> -1 Then
        Set oNode = moTitleList.ItemPhysical(lstTitles.ListIndex)
        txtCode.Text = oNode.Value.Text
    End If
End Sub

Private Sub lstTitles_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oNode As clsNode
    Dim oInfo As clsInfo
    Dim lListIndex As Long
    Dim lLogicalKey As Long
    
    If lstTitles.ListIndex <> -1 Then
        Set oNode = moTitleList.ItemPhysical(lstTitles.ListIndex)
    End If
    
    Select Case KeyCode
        Case 46 'delete
            If lstTitles.ListIndex <> -1 Then
                lListIndex = lstTitles.ListIndex
                moTitleList.RemovePhysical lstTitles.ListIndex
                lstTitles.RemoveItem lstTitles.ListIndex
                If lListIndex > 1 Then
                    lstTitles.ListIndex = lListIndex - 1
                Else
                    If lstTitles.ListCount > 0 Then
                        lstTitles.ListIndex = 0
                    Else
                        txtCode.Text = ""
                    End If
                End If
            End If
        Case 13 'enter
            frmEnterText.txtNote.Text = oNode.Value.Title
            frmEnterText.Show vbModal
            oNode.Value.Title = frmEnterText.txtNote.Text
            lstTitles.List(lstTitles.ListIndex) = frmEnterText.txtNote.Text
            KeyCode = 0
        Case 78 'new
            frmEnterText.txtNote.Text = ""
            frmEnterText.Show vbModal
            Set oNode = moTitleList.AddNew
            Set oInfo = New clsInfo
            oInfo.Title = frmEnterText.txtNote.Text
            Set oNode.Value = oInfo
            lstTitles.AddItem frmEnterText.txtNote.Text
            lstTitles.ListIndex = oNode.PhysicalKey
        Case 40 'down
            If Shift <> 0 Then
                If Not oNode Is Nothing Then
                    lLogicalKey = oNode.LogicalKey
                    lListIndex = lstTitles.ListIndex
                    If lListIndex < lstTitles.ListCount - 1 Then
                        moTitleList.Move lstTitles.ListIndex, lstTitles.ListIndex + 1
                        lstTitles.RemoveItem lstTitles.ListIndex
                        Set oNode = moTitleList.ItemLogical(lLogicalKey)
                        lstTitles.AddItem oNode.Value.Title, lListIndex + 1
                        lstTitles.ListIndex = lListIndex
                    End If
                End If
            End If
        Case 38 'up
            If Shift <> 0 Then
                If Not oNode Is Nothing Then
                    lLogicalKey = oNode.LogicalKey
                    lListIndex = lstTitles.ListIndex
                    If lListIndex > 0 Then
                        moTitleList.Move lstTitles.ListIndex, lstTitles.ListIndex - 1
                        lstTitles.RemoveItem lstTitles.ListIndex
                        Set oNode = moTitleList.ItemLogical(lLogicalKey)
                        lstTitles.AddItem oNode.Value.Title, lListIndex - 1
                        lstTitles.ListIndex = lListIndex
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtCode_Change()
    Dim oNode As clsNode
    
    If lstTitles.ListIndex <> -1 Then
        Set oNode = moTitleList.ItemPhysical(lstTitles.ListIndex)
        oNode.Value.Text = txtCode.Text
    End If
End Sub
