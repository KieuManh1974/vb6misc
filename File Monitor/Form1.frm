VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Monitor"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCopyFile 
      Height          =   285
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   2760
      Width           =   7935
   End
   Begin VB.Timer tmrPoll 
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   9135
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   9135
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
   Begin VB.ListBox lstFiles 
      Height          =   2085
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "Copy File From:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vCopyFiles As Variant

Private Sub cmdAdd_Click()
    Dim oFSO As New FileSystemObject
    
    If oFSO.FileExists(txtPath.Text) Then
        lstFiles.AddItem txtPath.Text
    Else
        txtPath.ForeColor = vbRed
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lIndex As Long
    
    If lstFiles.ListIndex <> -1 Then
        For lIndex = lstFiles.ListIndex To UBound(vCopyFiles) - 1
            vCopyFiles(lIndex) = vCopyFiles(lIndex + 1)
        Next
        lstFiles.RemoveItem lstFiles.ListIndex
        If UBound(vCopyFiles) = 0 Then
            vCopyFiles = Array()
            txtCopyFile.Text = ""
            txtPath.Text = ""
        Else
            ReDim Preserve vCopyFiles(UBound(vCopyFiles) - 1) As Variant
            txtCopyFile.Text = vCopyFiles(lstFiles.ListIndex)
        End If
    End If
End Sub

Private Sub Form_Load()
    vCopyFiles = Array()
End Sub

Private Sub lstFiles_Click()
    If lstFiles.ListIndex <> -1 Then
        txtPath.Text = lstFiles.List(lstFiles.ListIndex)
        txtCopyFile.Text = vCopyFiles(lstFiles.ListIndex)
    End If
End Sub

Private Sub tmrPoll_Timer()
    Dim lIndex As Long
    Dim oFSO As New FileSystemObject
    Dim oFile As File
    
    On Error Resume Next
    For lIndex = 1 To lstFiles.ListCount
        If lstFiles.Selected(lIndex - 1) = True Then
            If Not oFSO.FileExists(lstFiles.List(lIndex - 1)) Then
                Me.SetFocus
                MsgBox "File: " & lstFiles.List(lIndex - 1) & " is free."
                lstFiles.Selected(lIndex - 1) = False
            Else
                Set oFile = oFSO.GetFile(lstFiles.List(lIndex - 1))
                oFile.Attributes = oFile.Attributes Xor ReadOnly
                If Err.Number = 0 Then
                    Me.SetFocus

                    lstFiles.Selected(lIndex - 1) = False
                    oFile.Attributes = oFile.Attributes Xor ReadOnly
                    
                    If oFSO.FileExists(vCopyFiles(lIndex - 1)) Then
                        If oFSO.GetFile(vCopyFiles(lIndex - 1)).Name = oFSO.GetFile(lstFiles.List(lIndex - 1)).Name Then
                            oFSO.CopyFile vCopyFiles(lIndex - 1), lstFiles.List(lIndex - 1), True
                            MsgBox "File: " & lstFiles.List(lIndex - 1) & " is updated from: " & vbCrLf & vCopyFiles(lIndex - 1) & "(" & Now & ")"
                        Else
                            MsgBox "File: " & lstFiles.List(lIndex - 1) & " is free."
                        End If
                    Else
                        MsgBox "File: " & lstFiles.List(lIndex - 1) & " is free."
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Sub txtCopyFile_Change()
    txtCopyFile.ForeColor = vbBlack
    If lstFiles.ListIndex <> -1 Then
        vCopyFiles(lstFiles.ListIndex) = txtCopyFile.Text
    End If
End Sub

Private Sub txtCopyFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oFSO As New FileSystemObject
    
    If lstFiles.ListIndex <> -1 Then
        If oFSO.FileExists(Data.Files(1)) Then
            txtCopyFile.Text = oFSO.GetFile(Data.Files(1)).Path
        End If
        txtCopyFile.ForeColor = vbBlack
    
        vCopyFiles(lstFiles.ListIndex) = txtCopyFile.Text
    End If
End Sub

Private Sub txtPath_Change()
    txtPath.ForeColor = vbBlack
End Sub

Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtPath.Text = Data.Files(1)
    txtPath.ForeColor = vbBlack
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oFSO As New FileSystemObject
    
    If oFSO.FileExists(Data.Files(1)) Then
        lstFiles.AddItem Data.Files(1)
        ReDim Preserve vCopyFiles(UBound(vCopyFiles) + 1) As Variant
        vCopyFiles(UBound(vCopyFiles)) = ""
    End If
End Sub

