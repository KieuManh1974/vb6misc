VERSION 5.00
Begin VB.Form frmPriorityList 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Priority List"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstNotes 
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
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuNewNote 
         Caption         =   "&New Note"
      End
      Begin VB.Menu mnuRemoveNote 
         Caption         =   "&Remove Note"
      End
      Begin VB.Menu mnuEditNote 
         Caption         =   "&Edit Note"
      End
   End
End
Attribute VB_Name = "frmPriorityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sTemp As String
    Dim iListIndex As Integer
    
    Select Case KeyCode
        Case 38 ' Up
            If lstNotes.ListIndex <> -1 Then
                iListIndex = lstNotes.ListIndex
                If iListIndex <> 0 Then
                    sTemp = lstNotes.List(iListIndex)
                    lstNotes.RemoveItem (iListIndex)
                    lstNotes.AddItem sTemp, iListIndex - 1
                End If
            End If
        Case 40 ' Down
            If lstNotes.ListIndex <> -1 Then
                iListIndex = lstNotes.ListIndex
                If iListIndex <> lstNotes.ListCount - 1 Then
                    sTemp = lstNotes.List(iListIndex)
                    lstNotes.RemoveItem (iListIndex)
                    lstNotes.AddItem sTemp, iListIndex + 1
                    lstNotes.ListIndex = iListIndex
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim sNotes As String
    Dim vNotes As Variant
    Dim iIndex As Integer
    Dim vNote As Variant
    
    Me.Width = GetSetting("PriorityList", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("PriorityList", "Dimensions", "Height", Me.Height)
    Me.Left = GetSetting("PriorityList", "Position", "Left", Me.Left)
    Me.Top = GetSetting("PriorityList", "Position", "Top", Me.Top)
    
    sNotes = GetSetting("PriorityList", "Notes", "Notes", "")
    vNotes = Split(sNotes, vbCr)
    For Each vNote In vNotes
        If vNote <> "" Then
            lstNotes.AddItem vNote
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmEnterText
End Sub

Private Sub Form_Resize()
    With lstNotes
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sNotes As String
    Dim iIndex As Integer
    
    For iIndex = 0 To lstNotes.ListCount - 1
        sNotes = sNotes & lstNotes.List(iIndex) & vbCr
    Next
    
    SaveSetting "PriorityList", "Dimensions", "Width", Me.Width
    SaveSetting "PriorityList", "Dimensions", "Height", Me.Height
    SaveSetting "PriorityList", "Position", "Left", Me.Left
    SaveSetting "PriorityList", "Position", "Top", Me.Top
    SaveSetting "PriorityList", "Notes", "Notes", sNotes
End Sub

Private Sub lstNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuActions
    End If
End Sub

Private Sub mnuEditNote_Click()
    If lstNotes.ListIndex <> -1 Then
        frmEnterText.txtNote.Text = lstNotes.List(lstNotes.ListIndex)
        frmEnterText.Show vbModal
        lstNotes.List(lstNotes.ListIndex) = frmEnterText.txtNote.Text
    End If
End Sub

Private Sub mnuNewNote_Click()
    Dim iIndex As Integer
    
    frmEnterText.txtNote = ""
    frmEnterText.Show vbModal

    If lstNotes.ListIndex <> -1 Then
        lstNotes.AddItem frmEnterText.txtNote.Text, lstNotes.ListIndex
    Else
        lstNotes.AddItem frmEnterText.txtNote.Text
    End If
End Sub

Private Sub mnuRemoveNote_Click()
    Dim iIndex As Integer
    
    If lstNotes.ListIndex <> -1 Then
        lstNotes.RemoveItem (lstNotes.ListIndex)
    End If
End Sub
