VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Music Tester"
   ClientHeight    =   3270
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6855
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuTrebleClef 
         Caption         =   "Treble Clef"
      End
      Begin VB.Menu mnuBassClef 
         Caption         =   "Bass Clef"
      End
      Begin VB.Menu mnuNoteCount 
         Caption         =   "Note Count"
         WindowList      =   -1  'True
         Begin VB.Menu mnuNoteCount1 
            Caption         =   "1"
         End
         Begin VB.Menu mnuNoteCount2 
            Caption         =   "2"
         End
         Begin VB.Menu mnuNoteCount3 
            Caption         =   "3"
         End
         Begin VB.Menu mnuNoteCount4 
            Caption         =   "4"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 17 Bass Clef (D)
' 29 Trebble Clef (B)

Private Const mlOffsetX As Long = 30
Private Const mlOffsetY As Long = 50
Private Const mlStaffSpacing As Long = 10
Private Const mlNoteSpacing As Long = 37
Private Const mlStaffWidth As Long = 400
Private Const mlLedgerWidth As Long = 34

Private msNoteNames As String
Private mlNumberOfNotes As Long
Private mlStaveOffset As Long

Private Sub Form_Activate()
    mlStaveOffset = 17
    mlNumberOfNotes = 2
    GenerateSequence
End Sub

Private Sub GenerateSequence()
    Dim vNotes As Variant
    
    Text1.Text = ""
    vNotes = GenerateNotes(mlNumberOfNotes, mlStaveOffset)
    msNoteNames = NoteNames(vNotes)
    DrawNotes vNotes, mlStaveOffset

End Sub

Public Function GenerateNotes(lNumber As Long, lStaveOffset As Long) As Variant
    Dim lIndex As Long
    Dim vNotes As Variant
    
    Randomize
    vNotes = Array()
    ReDim vNotes(lNumber - 1)
    
    For lIndex = 0 To lNumber - 1
        vNotes(lIndex) = lStaveOffset + Int(Rnd * 23) - 11
    Next
    GenerateNotes = vNotes
End Function

Private Function NoteNames(vNotes As Variant) As String
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vNotes)
        NoteNames = NoteNames & Mid$("ABCDEFG", (vNotes(lIndex) Mod 7) + 1, 1)
    Next
    
End Function

Public Sub DrawNotes(vNotes As Variant, lStaveOffset As Long)
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    Cls
    
    For lIndex = 0 To 4
        Line (mlOffsetX, lIndex * mlStaffSpacing + mlOffsetY)-Step(mlStaffWidth, 0), vbBlack
    Next
    
    Me.FillStyle = vbSolid
    For lIndex = 0 To UBound(vNotes)
        Me.DrawWidth = 1
        Me.Circle (mlNoteSpacing * 2 + lIndex * mlNoteSpacing, mlOffsetY - (vNotes(lIndex) - lStaveOffset - 4) * mlStaffSpacing \ 2), 10, vbBlack, , , 0.45
        Me.DrawWidth = 2
        If (vNotes(lIndex) - lStaveOffset) >= 0 Then
            Me.Line (mlNoteSpacing * 2 + lIndex * mlNoteSpacing - 10, mlOffsetY - (vNotes(lIndex) - lStaveOffset - 4) * mlStaffSpacing \ 2)-Step(0, 30), vbBlack
        Else
            Me.Line (mlNoteSpacing * 2 + lIndex * mlNoteSpacing + 10, mlOffsetY - (vNotes(lIndex) - lStaveOffset - 4) * mlStaffSpacing \ 2)-Step(0, -30), vbBlack
        End If
        Select Case vNotes(lIndex) - lStaveOffset
            Case Is <= -6
                For lIndex2 = -6 To (vNotes(lIndex) - lStaveOffset) Step -2
                    Me.Line (mlNoteSpacing * 2 + lIndex * mlNoteSpacing - mlLedgerWidth \ 2, mlOffsetY - (lIndex2 - 4) * mlStaffSpacing \ 2)-Step(mlLedgerWidth, 0), vbBlack
                Next
            Case Is >= 6
                For lIndex2 = 6 To (vNotes(lIndex) - lStaveOffset) Step 2
                    Me.Line (mlNoteSpacing * 2 + lIndex * mlNoteSpacing - mlLedgerWidth \ 2, mlOffsetY - (lIndex2 - 4) * mlStaffSpacing \ 2)-Step(mlLedgerWidth, 0), vbBlack
                Next
        End Select
    Next
    
End Sub

Private Sub mnuTrebleClef_Click()
    mnuTrebleClef.Checked = vbChecked
    mnuBassClef.Checked = Not mnuTrebleClef.Checked
    mlStaveOffset = 29
    GenerateSequence
End Sub

Private Sub mnuBassClef_Click()
    mnuBassClef.Checked = vbChecked
    mnuTrebleClef.Checked = vbUnchecked
    mlStaveOffset = 17
    GenerateSequence
End Sub

Private Sub mnuNoteCount1_Click()
    mnuNoteCount1.Checked = vbChecked
    mnuNoteCount2.Checked = vbUnchecked
    mnuNoteCount3.Checked = vbUnchecked
    mnuNoteCount4.Checked = vbUnchecked
    mlNumberOfNotes = 1
    GenerateSequence
End Sub

Private Sub mnuNoteCount2_Click()
    mnuNoteCount2.Checked = vbChecked
    mnuNoteCount1.Checked = vbUnchecked
    mnuNoteCount3.Checked = vbUnchecked
    mnuNoteCount4.Checked = vbUnchecked
    mlNumberOfNotes = 2
    GenerateSequence
End Sub

Private Sub mnuNoteCount3_Click()
    mnuNoteCount3.Checked = vbChecked
    mnuNoteCount1.Checked = vbUnchecked
    mnuNoteCount2.Checked = vbUnchecked
    mnuNoteCount4.Checked = vbUnchecked
    mlNumberOfNotes = 3
    GenerateSequence
End Sub

Private Sub mnuNoteCount4_Click()
    mnuNoteCount4.Checked = vbChecked
    mnuNoteCount1.Checked = vbUnchecked
    mnuNoteCount2.Checked = vbUnchecked
    mnuNoteCount3.Checked = vbUnchecked
    mlNumberOfNotes = 4
    GenerateSequence
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    Text1.ForeColor = vbBlack
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii >= 32 Then
        If KeyAscii > Asc("G") Or KeyAscii < Asc("A") Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        If Text1.Text = msNoteNames Then
            GenerateSequence
        Else
            Text1.ForeColor = vbRed
        End If
    End If
End Sub
