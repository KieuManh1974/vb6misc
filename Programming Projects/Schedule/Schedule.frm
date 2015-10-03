VERSION 5.00
Begin VB.Form frmSchedule 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Schedule"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   720
      ScaleHeight     =   735
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuNewTask 
         Caption         =   "&New Note"
      End
      Begin VB.Menu mnuRemoveTask 
         Caption         =   "&Remove Task"
      End
      Begin VB.Menu mnuEditTask 
         Caption         =   "&Edit Task"
      End
      Begin VB.Menu mnuCompleteTask 
         Caption         =   "&Complete Task"
      End
      Begin VB.Menu mnuHoldTask 
         Caption         =   "&Hold Task"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "On Top"
      End
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

Private moNotes As New clsRepresentation
Private moCopyTask As clsHierarchy

Private mnHeight As Single
Private mlSelection As Long

Private Const HilightColour As Long = &H80C0FF
Private Const CompletedColour As Long = &H808080
Private Const HoldColour As Long = &H80&

Private Sub Form_DblClick()
    mnuEditTask_Click
End Sub

Private Property Let OnTop(bOnTop As Boolean)
    mnuOnTop.Checked = bOnTop
    If bOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    Else
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End If
End Property

Private Property Get OnTop() As Boolean
    OnTop = mnuOnTop.Checked
End Property

Private Sub Form_Load()
    Me.Width = GetSetting("PriorityList", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("PriorityList", "Dimensions", "Height", Me.Height)
    Me.Left = GetSetting("PriorityList", "Position", "Left", Me.Left)
    Me.Top = GetSetting("PriorityList", "Position", "Top", Me.Top)
    Me.Font.Size = GetSetting("PriorityList", "Font", "Size", Me.Font.Size)
    OnTop = GetSetting("PriorityList", "Position", "OnTop", False)
        
    moNotes.LoadNotes
    
    mnHeight = TextHeight("X")
    mlSelection = -1
    ShowNotes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "PriorityList", "Dimensions", "Width", Me.Width
    SaveSetting "PriorityList", "Dimensions", "Height", Me.Height
    SaveSetting "PriorityList", "Position", "Left", Me.Left
    SaveSetting "PriorityList", "Position", "Top", Me.Top
    SaveSetting "PriorityList", "Font", "Size", Me.Font.Size
    SaveSetting "PriorityList", "Position", "OnTop", OnTop
    moNotes.SaveNotes
End Sub

Private Sub DisplayTask(ByVal lIndex As Long, Optional ByVal bSelect As Boolean)
    Dim nWidth As Single
    Dim nTop As Single
    Dim oHierarchy As clsHierarchy
    Dim sDescription As String
    
    Set oHierarchy = moNotes.Displaylist.List(lIndex + 1)
    sDescription = IIf(Not oHierarchy.Expanded And oHierarchy.Children.Count > 0, "+", "") & oHierarchy.Task.Description
    
    nTop = lIndex * mnHeight
    nWidth = TextWidth(sDescription) + 200 * oHierarchy.Level
    
    If Not bSelect Then
        Me.Line (0, nTop)-Step(nWidth, mnHeight), BackColor, BF
    Else
        Me.Line (0, nTop)-Step(nWidth, mnHeight), HilightColour, BF
    End If
    
    Me.CurrentX = 200 * oHierarchy.Level
    Me.CurrentY = nTop
    
    If oHierarchy.Task.Status = TaskStatuses.Completed Then
        ForeColor = CompletedColour
    ElseIf oHierarchy.Task.Status = TaskStatuses.Hold Then
        ForeColor = HoldColour
    ElseIf oHierarchy.Task.Status = TaskStatuses.Active Then
        ForeColor = vbBlack
    End If
    Print sDescription
End Sub


Private Sub ShowNotes(Optional ByVal lListIndex As Long = -1)
    Dim oHierarchy As clsHierarchy
    Dim lIndex As Long
    
    Cls
    For lIndex = 1 To moNotes.Displaylist.List.Count
        DisplayTask lIndex - 1
    Next
    
    Selection = lListIndex
    moNotes.SaveNotes
End Sub

Private Property Let Selection(ByVal lIndex As Long)
    If mlSelection <> -1 And mlSelection < moNotes.Displaylist.List.Count Then
        DisplayTask mlSelection
    End If

    If lIndex < -1 Then
        mlSelection = moNotes.Displaylist.List.Count - 1
    ElseIf lIndex >= moNotes.Displaylist.List.Count Then
        mlSelection = -1
    Else
        mlSelection = lIndex
    End If
    If mlSelection > -1 And mlSelection < moNotes.Displaylist.List.Count Then
        DisplayTask mlSelection, True
    End If
End Property

Private Property Get Selection() As Long
    Selection = mlSelection
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oSelectedTask As clsTask
    Dim lListIndex As Long
    Dim lNewListIndex As Long
    
    lListIndex = mlSelection
    If lListIndex <> -1 Then
        Select Case KeyCode
            Case 9 ' Unselect
                Selection = -1
            Case 107, 187 ' Expand
                If Shift = 0 Then
                    moNotes.Displaylist.List(lListIndex + 1).Expanded = True
                    moNotes.RefreshDisplayListFromList
                    ShowNotes lListIndex
                    KeyCode = 0
                End If
            Case 109, 189 ' Collapse
                If Shift = 0 Then
                    moNotes.Displaylist.List(lListIndex + 1).Expanded = False
                    moNotes.RefreshDisplayListFromList
                    ShowNotes lListIndex
                    KeyCode = 0
                End If
            Case 67 ' Complete
                mnuCompleteTask_Click
                KeyCode = 0
            Case 72 ' Hold
                mnuHoldTask_Click
                KeyCode = 0
            Case 69, 13 ' Edit
                mnuEditTask_Click
                KeyCode = 0
            Case 38 ' Up
                If Shift = 1 Then
                    lNewListIndex = moNotes.MoveTaskUp(moNotes.Displaylist.List(lListIndex + 1))
                    If lNewListIndex > -1 Then
                        ShowNotes lNewListIndex
                    End If
                    KeyCode = 0
                ElseIf Shift = 2 Then
                    lNewListIndex = moNotes.Displaylist.List(lListIndex + 1).Parent.FindIdentifierIndex(moNotes.Displaylist.List(lListIndex + 1).Identifier) - 1
                    If lNewListIndex = 0 Then
                        lNewListIndex = 1
                    End If
                    lNewListIndex = moNotes.Displaylist.FindIdentifierIndex(moNotes.Displaylist.List(lListIndex + 1).Parent.Children(lNewListIndex).Identifier)
                    Selection = lNewListIndex
                End If
            Case 40 ' Down
                If Shift = 1 Then
                    lNewListIndex = moNotes.MoveTaskDown(moNotes.Displaylist.List(lListIndex + 1))
                    If lNewListIndex > -1 Then
                        ShowNotes lNewListIndex
                    End If
                    KeyCode = 0
                ElseIf Shift = 2 Then
                    lNewListIndex = moNotes.Displaylist.List(lListIndex + 1).Parent.FindIdentifierIndex(moNotes.Displaylist.List(lListIndex + 1).Identifier) + 1
                    If lNewListIndex > moNotes.Displaylist.List(lListIndex + 1).Parent.Children.Count Then
                        lNewListIndex = moNotes.Displaylist.List(lListIndex + 1).Parent.Children.Count
                    End If
                    lNewListIndex = moNotes.Displaylist.FindIdentifierIndex(moNotes.Displaylist.List(lListIndex + 1).Parent.Children(lNewListIndex).Identifier)
                    Selection = lNewListIndex
                End If
            Case 39 ' Right
            Case 37 ' Left
            Case 46, 8 ' Delete
                If Shift <> 0 Then
                    mnuRemoveTask_Click
                    KeyCode = 0
                End If
            Case 88 ' Cut
                Set moCopyTask = moNotes.Displaylist.List(lListIndex + 1)
        End Select
    End If
    Select Case KeyCode
        Case 78 ' New
            mnuNewTask_Click
            KeyCode = 0
        Case 86 ' Paste
            If Not moCopyTask Is Nothing Then
                If lListIndex > -1 Then
                    ShowNotes moNotes.MoveTask(moCopyTask, moNotes.Displaylist.List(lListIndex + 1))
                Else
                    ShowNotes moNotes.MoveTask(moCopyTask, moNotes.Tree.Top)
                End If
                KeyCode = 0
            End If
        Case 38 ' Up
            If Shift = 0 Then
                Selection = Selection - 1
            End If
        Case 40 ' Down
            If Shift = 0 Then
                Selection = Selection + 1
            End If
        Case 107, 187 ' Increase font
            Dim nDelta As Single
            Dim nOldFontSize As Single
            
            If Shift = 2 Then
                nDelta = 0.1
                nOldFontSize = Font.Size
                While Font.Size = nOldFontSize
                    Font.Size = Font.Size + nDelta
                    nDelta = nDelta + 0.1
                Wend
                mnHeight = TextHeight("X")
                ShowNotes Selection
                KeyCode = 0
            End If
        Case 109, 189 ' Increase font
            If Shift = 2 Then
                nDelta = 0.1
                nOldFontSize = Font.Size
                While Font.Size = nOldFontSize
                    Font.Size = Font.Size - nDelta
                    nDelta = nDelta + 0.1
                Wend
                mnHeight = TextHeight("X")
                ShowNotes Selection
                KeyCode = 0
            End If
    End Select
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lIndex As Long
    
    lIndex = Int(y / mnHeight)
    If lIndex < moNotes.Displaylist.List.Count Then
        Selection = lIndex
    Else
        Selection = -1
    End If
    
    If Button = vbRightButton Then
        Me.PopupMenu mnuActions
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmEnterText
End Sub


Private Sub mnuEditTask_Click()
    If mlSelection <> -1 Then
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    
        frmEnterText.txtNote.Text = moNotes.Displaylist.List(mlSelection + 1).Task.Description
        frmEnterText.Show vbModal
            
        If mnuOnTop.Checked Then
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        End If
        If Trim$(frmEnterText.txtNote.Text) <> "" Then
            moNotes.Displaylist.List(mlSelection + 1).Task.Description = Trim$(frmEnterText.txtNote.Text)
            ShowNotes mlSelection
        End If
    End If
End Sub

Private Sub mnuCompleteTask_Click()
    If mlSelection <> -1 Then
        If moNotes.Displaylist.List(mlSelection + 1).Task.Status <> TaskStatuses.Completed Then
            moNotes.Displaylist.List(mlSelection + 1).Task.Status = TaskStatuses.Completed
        Else
            moNotes.Displaylist.List(mlSelection + 1).Task.Status = TaskStatuses.Active
        End If
        
        ShowNotes mlSelection
    End If
End Sub

Private Sub mnuHoldTask_Click()
    If mlSelection <> -1 Then
        If moNotes.Displaylist.List(mlSelection + 1).Task.Status <> TaskStatuses.Hold Then
            moNotes.Displaylist.List(mlSelection + 1).Task.Status = TaskStatuses.Hold
        Else
            moNotes.Displaylist.List(mlSelection + 1).Task.Status = TaskStatuses.Active
        End If
        
        ShowNotes mlSelection
    End If
End Sub


Private Sub mnuNewTask_Click()
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    
    frmEnterText.txtNote.Text = ""
    frmEnterText.Show vbModal
    
    If mnuOnTop.Checked Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    End If
    
    If Trim$(frmEnterText.txtNote.Text) <> "" Then
        If mlSelection = -1 Then
            ShowNotes moNotes.NewTask(moNotes.Tree.Top, Trim$(frmEnterText.txtNote.Text))
        Else
            ShowNotes moNotes.NewTask(moNotes.Displaylist.List(mlSelection + 1), Trim$(frmEnterText.txtNote.Text))
        End If
    End If
End Sub

Private Sub mnuOnTop_Click()
    OnTop = Not OnTop
End Sub

Private Sub mnuRemoveTask_Click()
    If mlSelection <> -1 Then
        ShowNotes moNotes.RemoveTask(moNotes.Displaylist.List(mlSelection + 1))
    End If
End Sub
