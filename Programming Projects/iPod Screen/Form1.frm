VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "iPod Screen"
   ClientHeight    =   1980
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   2640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   78.135
   ScaleMode       =   0  'User
   ScaleWidth      =   125.409
   StartUpPosition =   3  'Windows Default
   Begin DHTMLEDLibCtl.DHTMLEdit htmPanel 
      CausesValidation=   0   'False
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2655
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   0
      Scrollbars      =   0   'False
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   0
      ScaleHeight     =   1980
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlFileIndex As Long

Private Sub Form_Load()
    Dim oFSO As New FileSystemObject
    Dim sHTML As String
    
    sHTML = oFSO.OpenTextFile(App.Path & "\info.htm").ReadAll
    
    htmPanel.DocumentHTML = sHTML
End Sub

Private Sub htmPanel_onkeypress()
    Unload Me
End Sub

Private Sub htmPanel_onmousedown()
    keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
    Clipboard.Clear
    
    DoEvents
    Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    While Dir(App.Path & "\ipod" & mlFileIndex & ".bmp") <> ""
        mlFileIndex = mlFileIndex + 1
    Wend
    SavePicture Picture1.Picture, App.Path & "\ipod" & mlFileIndex & ".bmp"
End Sub
