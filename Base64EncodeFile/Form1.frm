VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   4440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "Load"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    txtFile.Width = Me.Width - 150
    txtFile.Height = Me.Height - 50
End Sub

Private Sub mnuLoad_Click()
    On Error GoTo mnuLoad_ClickExit
 
    cdgFile.CancelError = True
    'comFile.FileName = sPath
    cdgFile.Filter = "*.*"
    cdgFile.ShowOpen

    txtFile.Text = OpenFileAndEncode(cdgFile.FileName)
    Clipboard.Clear
    Clipboard.SetText txtFile.Text, vbCFText
    
mnuLoad_ClickExit:

End Sub

Private Sub mnuSave_Click()
    On Error GoTo mnuSave_ClickExit
     
    cdgFile.CancelError = True
    'comFile.FileName = sPath
    cdgFile.Filter = "*.*"
    cdgFile.ShowSave
    
    DecodeAndSaveFile cdgFile.FileName, txtFile.Text
    
mnuSave_ClickExit:
End Sub
