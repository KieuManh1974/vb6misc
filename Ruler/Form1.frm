VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ruler"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScreenHeightCM 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cboUnits 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Text            =   "Centimetres"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mnScale As Single
Private mnScreenHeightCM As Single

Private Sub cboUnits_Click()
    Select Case cboUnits.ListIndex
        Case 0
            mnScale = (Screen.Height \ Screen.TwipsPerPixelY) / mnScreenHeightCM
            DrawRuler
        Case 1
            mnScale = 2.54 * (Screen.Height \ Screen.TwipsPerPixelY) / mnScreenHeightCM
            DrawRuler
    End Select
    SaveSetting "Ruler", "Units", "Unit", cboUnits.ListIndex
End Sub

Private Sub Form_Load()
    mnScreenHeightCM = GetSetting("Ruler", "Units", "ScreenHeightCM", 21.4)
    txtScreenHeightCM.Text = mnScreenHeightCM
    cboUnits.ListIndex = GetSetting("Ruler", "Units", "Unit", 0)
    cboUnits_Click
End Sub

Private Sub Form_Resize()
    DrawRuler
End Sub

Private Sub DrawRuler()
    Dim nX As Single
    Dim nY As Single
    Dim nPad As Single
    Dim lIndex As Long
    
    Cls
    nPad = 50
    
   For nX = nPad To Me.Width \ Screen.TwipsPerPixelX - nPad Step mnScale
        Me.Line (nX, nPad)-Step(0, Me.Height \ Screen.TwipsPerPixelY - nPad * 2)
        Me.CurrentX = nX - 7
        Me.CurrentY = nPad - 14
        Print lIndex
        lIndex = lIndex + 1
    Next
    
    lIndex = 0
    For nY = nPad To Me.Height \ Screen.TwipsPerPixelY - nPad Step mnScale
        Me.Line (nPad, nY)-Step(Me.Width \ Screen.TwipsPerPixelX - nPad * 2, 0)
        Me.CurrentX = nPad - 20
        Me.CurrentY = nY - 7
        Print lIndex
        lIndex = lIndex + 1
    Next
End Sub


Private Sub txtScreenHeightCM_LostFocus()
    If IsNumeric(txtScreenHeightCM.Text) Then
        mnScreenHeightCM = Val(txtScreenHeightCM.Text)
        SaveSetting "Ruler", "Units", "ScreenHeightCM", txtScreenHeightCM.Text
        cboUnits_Click
    End If
End Sub
