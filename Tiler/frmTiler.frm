VERSION 5.00
Begin VB.Form frmTiler 
   BackColor       =   &H00000000&
   Caption         =   "Tiler"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrZoom 
      Height          =   255
      Left            =   480
      Max             =   340
      Min             =   50
      TabIndex        =   15
      Top             =   5040
      Value           =   50
      Width           =   1695
   End
   Begin VB.HScrollBar scrBlue 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.HScrollBar scrGreen 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   10
      Top             =   4200
      Width           =   1695
   End
   Begin VB.HScrollBar scrRed 
      Height          =   255
      Left            =   480
      Max             =   255
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox pctTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   14775
      Left            =   2280
      ScaleHeight     =   981
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1117
      TabIndex        =   8
      Top             =   0
      Width           =   16815
   End
   Begin VB.ComboBox cboTile 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox cboTile 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox cboTile 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtPattern 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmTiler.frx":0000
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      ItemData        =   "frmTiler.frx":0002
      Left            =   120
      List            =   "frmTiler.frx":000C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pattern"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moPics() As StdPicture
Private mnWidths() As Single

Private nScale As Single
Private moParser As IParseObject

Private Sub cboStyle_Click()
    cmdGo_Click
End Sub

Private Sub cboTile_Click(Index As Integer)
    cmdGo_Click
End Sub

Private Sub cmdGo_Click()
    Dim vPattern As Variant
    
    vPattern = SplitPattern
    RenderPattern vPattern, cboStyle.ListIndex = 1
End Sub

Private Function SplitPattern() As Variant
    Dim vCourses As Variant
    Dim vCourse As Variant
    Dim vLine As Variant
    Dim vPattern As Variant
    
    vPattern = Array()
    
    vCourses = Split(txtPattern.Text, "/")
    For Each vCourse In vCourses
        vLine = Split(vCourse, ",")
        ReDim Preserve vPattern(UBound(vPattern) + 1)
        vPattern(UBound(vPattern)) = vLine
    Next
    SplitPattern = vPattern
End Function

Private Sub RenderPattern(vPattern As Variant, bStaggered As Boolean)
    Dim lX As Long
    Dim lY As Long
    Dim oPic As StdPicture
    Dim lPicIndex As Long
    Dim nOffset As Single
    Dim lYMax As Long
    Dim lXMax As Long
    Dim nAspect As Single
    Dim nY As Single
    Dim nX As Single
    Dim nMaxHeight As Single
    
    lXMax = pctTiles.Width / Screen.TwipsPerPixelX / nScale + 1
    lYMax = pctTiles.Height / Screen.TwipsPerPixelY / nScale + 1
    pctTiles.Cls
    For lY = 0 To lYMax / 2
        If bStaggered Then
            If lY Mod 2 = 0 Then
                nOffset = -nScale / 2
            Else
                nOffset = 0
            End If
        End If
        nMaxHeight = 0
        nX = 0
        For lX = 0 To lXMax
            lPicIndex = cboTile(vPattern(lY Mod (UBound(vPattern) + 1))(lX Mod (1 + UBound(vPattern(lY Mod (UBound(vPattern) + 1))))) - 1).ListIndex
            nAspect = moPics(lPicIndex).Height / moPics(lPicIndex).Width
            If nAspect * nScale * mnWidths(lPicIndex) / 100 > nMaxHeight Then
                nMaxHeight = nAspect * nScale * mnWidths(lPicIndex) / 100
            End If
            pctTiles.PaintPicture moPics(lPicIndex), nOffset + nX, pctTiles.Height / Screen.TwipsPerPixelY / 3 + nY, nScale * mnWidths(lPicIndex) / 100, nScale * nAspect * mnWidths(lPicIndex) / 100
            nX = nX + nScale * mnWidths(lPicIndex) / 100
        Next
        nY = nY + nMaxHeight
    Next
End Sub

Private Sub Form_Load()
    Dim oFSO As New FileSystemObject
    Dim oFolder As Folder
    Dim oFile As File
    Dim sExt As String
    Dim lDot As Long
    Dim sName As String
    Dim lPicIndex As Long
    Dim oTree As ParseTree
    Dim lTileWidth As Long
    
    InitParser
    nScale = scrZoom.Value
    
    Set oFolder = oFSO.GetFolder(App.Path)
    For Each oFile In oFolder.Files
        lDot = InStrRev(oFile.Name, ".")
        sExt = LCase$(Mid$(oFile.Name, lDot + 1))
        Select Case sExt
            Case "jpg", "bmp", "gif"
                sName = Left$(oFile.Name, lDot)
                cboTile(0).AddItem sName
                cboTile(1).AddItem sName
                cboTile(2).AddItem sName
                ReDim Preserve moPics(lPicIndex)
                ReDim Preserve mnWidths(lPicIndex)
                Set moPics(lPicIndex) = LoadPicture(App.Path & "/" & oFile.Name)

                Stream.Text = oFile.Name
                Set oTree = New ParseTree
                If moParser.Parse(oTree) Then
                    mnWidths(lPicIndex) = oTree.Text
                Else
                    mnWidths(lPicIndex) = 100
                End If
                
                lPicIndex = lPicIndex + 1
        End Select
    Next
    cboTile(0).ListIndex = 0
    cboTile(1).ListIndex = 0
    cboTile(2).ListIndex = 0
    cboStyle.ListIndex = 0
End Sub

Private Sub InitParser()
    If Not SetNewDefinition("digits := repeat in '0'-'9' min 3 max 4; number := and [repeat skip until or digits,eos], digits;") Then
        MsgBox "Bad Def"
    End If
    Set moParser = ParserObjects("number")
End Sub

Private Sub scrBlue_Change()
    pctTiles.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
    cmdGo_Click
End Sub

Private Sub scrGreen_Change()
    pctTiles.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
    cmdGo_Click
End Sub

Private Sub scrRed_Change()
    pctTiles.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
    cmdGo_Click
End Sub

Private Sub scrZoom_Change()
    nScale = scrZoom.Value
    cmdGo_Click
End Sub
