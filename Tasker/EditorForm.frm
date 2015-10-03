VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EditorForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "[No Title]"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   3  'I-Beam
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6810
   ScaleWidth      =   9030
   Begin MSComCtl2.FlatScrollBar HorizScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      Arrows          =   65536
      LargeChange     =   64
      Max             =   2047
      Orientation     =   1179649
      SmallChange     =   10
   End
   Begin MSComCtl2.FlatScrollBar VertScroll 
      Height          =   1815
      Left            =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      Orientation     =   1179648
   End
   Begin MSComDlg.CommonDialog FilePicker 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Square 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open ..."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As ..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export to Visual Basic ..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "EditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type ColourInfo
    position As Integer
    colour As Long
    backcolour As Long
End Type

Private Type LineCode
    Text As String
    Codes() As ColourInfo
    CodesNotEmpty As Boolean
End Type

Private Type Coord
    X As Integer
    Y As Integer
End Type

Private Type AbsoluteCoord
    X As Single
    Y As Single
End Type

Private Cursor As Coord
Private OldCursor As Coord
Private LineHeight As Single
Private HiliteStart As Coord
Private HiliteOn As Boolean
Private Lines() As Task

Private mTopLineOffset As Integer
Private VisibleLines As Integer
Private Const LeftMargin As Single = 15 * 10
Private bInitialised As Boolean
Private mLeftSide As Single

Private mFileName As String

Private bCaretVisible As Boolean
Private bJustDown As Boolean

Private KeyLog() As String * 1
Private KeyLogCount As Integer

Private bSaveOnResize As Boolean
Private bTextChanged As Boolean

Private vAccented As Variant
Private vAccentedLetterIndex As String
Private vAccentIndex As String

Private RightPointer As StdPicture
Private DragPointer As StdPicture
Private DragCopyPointer As StdPicture
Private sDragText As String

Private Enum MouseModes
    Hilite = 0
    HiliteWord = 1
    Dragging = 2
    LineHilite = 3
    Default = 4
End Enum

Private Enum MouseDownModes
    mdmNotDown
    mdmHL
    mdmHLWord
    mdmHLLine
    mdmHLDragging
End Enum

Private Enum CursorModes
    cmNormal
    cmOverHL
    cmMargin
    cmDragMove
    cmDragCopy
End Enum

Private iMouseMode As MouseDownModes
Private iCursorMode As CursorModes
Private iLineHighlightIndex As Long


Private Property Let MouseMode(ByVal iNewMouseMode As MouseDownModes)
    iMouseMode = iNewMouseMode
End Property

Private Property Get MouseMode() As MouseDownModes
    MouseMode = iMouseMode
End Property

Private Property Let CursorMode(ByVal iNewCursorMode As CursorModes)
    Select Case iNewCursorMode
        Case cmNormal
            MousePointer = vbIbeam
        Case cmOverHL
            MousePointer = vbDefault
        Case cmMargin
            MousePointer = vbCustom
            MouseIcon = RightPointer
        Case cmDragMove
            MousePointer = vbCustom
            MouseIcon = DragPointer
        Case cmDragCopy
            MousePointer = vbCustom
            MouseIcon = DragCopyPointer
        End Select
        iCursorMode = iNewCursorMode
End Property

Private Property Get CursorMode() As CursorModes
    CursorMode = iCursorMode
End Property

Private Sub Form_Load()
    bSaveOnResize = True
    vAccented = Array(Array("¿¡¬√ƒ≈x", "‡·‚„‰Âx"), Array("»… xÀå∆", "ËÈÍxÎúÊ"), Array("ÃÕŒxœxx", "ÏÌÓxÔxx"), Array("“”‘’÷xxÿ", "ÚÛÙıˆxx¯"), Array("Ÿ⁄€x‹xxx", "˘˙˚x¸xxx"), Array("xxxxxxxx«", "xxxxxxxxÁ"), Array("xxx—xxxx", "xxxÒxxxx"))
    vAccentedLetterIndex = "aeioucn/"
    vAccentIndex = "`'^~:oa/,"
    
    Set RightPointer = LoadPicture(App.Path & "\NORMAL03.CUR")
    Set DragPointer = LoadPicture(App.Path & "\DRAGMOVE.CUR")
    Set DragCopyPointer = LoadPicture(App.Path & "\DRAGCOPY.CUR")
    
    LineHeight = Me.TextHeight("M")
End Sub

Private Sub Form_Resize()
    Static bFirstResizeWithFormOpen As Boolean
    
    VertScroll.Left = Me.ScaleWidth - VertScroll.Width
    VertScroll.Height = Me.ScaleHeight - HorizScroll.Height
    HorizScroll.Width = Me.ScaleWidth - VertScroll.Width
    HorizScroll.Top = Me.ScaleHeight - HorizScroll.Height
    Square.Left = VertScroll.Left
    Square.Top = HorizScroll.Top
    VisibleLines = 1 + Int((Me.ScaleHeight - HorizScroll.Height) / LineHeight)
    If VisibleLines < (UBound(Lines) - LBound(Lines)) Then
        VertScroll.Min = LBound(Lines) - 1
        VertScroll.Max = UBound(Lines) - 2
        VertScroll.SmallChange = 1
        VertScroll.LargeChange = VisibleLines
    End If
    TopLineOffset = mTopLineOffset
    
    If bSaveOnResize Then
        If Not bFirstResizeWithFormOpen Then
            InitialiseCursor
            bFirstResizeWithFormOpen = True
        End If
    End If
End Sub

Private Sub InitialiseConstants()
    LineHeight = Me.TextHeight("H")
End Sub

Private Sub InitialiseCursor()
    Cursor = position(1, 1)
    
    Dim h As Long
    h = Me.hwnd
    
    CreateCaret h, 0&, 0&, LineHeight / Screen.TwipsPerPixelY
    SetCaretPos LeftMargin \ 15, 0&
    ShowCaret h
    bCaretVisible = True
End Sub

Private Sub InitialiseLines(ByVal sFilename As String)
    Dim ts As TextStream
    Dim X As Integer

    ReDim Lines(1 To 1) As LineCode
    
    X = 1

    If sFilename <> "" Then
        With New FileSystemObject
            Me.Caption = .GetFile(sFilename).Name & " - PDL Editor"
            Set ts = .OpenTextFile(sFilename, ForReading)
            
            While Not ts.AtEndOfStream
                ReDim Preserve Lines(1 To X) As LineCode
                Lines(X).Text = ts.ReadLine
                X = X + 1
            Wend
        End With
        bTextChanged = True
        #If SyntaxCheck <> 0 Then
            AdjustAll
        #End If
        bTextChanged = False
    End If
    ReDim Preserve Lines(1 To X) As LineCode
    Lines(X).Text = ""
    
    TopLineOffset = 0
End Sub


Private Sub InitialiseLines2(ByVal sFilename As String)
    Dim ts As TextStream
    Dim X As Integer

    ReDim Lines(1 To 1) As LineCode
    
    X = 1

    If sFilename <> "" Then
        Open sFilename For Input As #1
        While Not EOF(1)
            ReDim Preserve Lines(1 To X) As LineCode
            Input #1, Lines(X).Text
            X = X + 1
        Wend
        #If SyntaxCheck <> 0 Then
            AdjustAll
        #End If
        
'        With New FileSystemObject
'            MsgBox sfilename
'
'            Me.Caption = .GetFile(sfilename).Name
'
'            Set ts = .OpenTextFile(sfilename, ForReading)
'
'            While Not ts.AtEndOfStream
'                ReDim Preserve Lines(1 To X) As LineCode
'                Lines(X).Text = ts.ReadLine
'                X = X + 1
'            Wend
'        End With
    End If
    ReDim Preserve Lines(1 To X) As LineCode
    Lines(X).Text = ""
    
    TopLineOffset = 0
End Sub


Public Sub LoadFile(sFilename As String)
    mFileName = sFilename

    InitialiseCursor
    InitialiseLines sFilename
    bInitialised = True
    Form_Resize
End Sub

Public Sub SaveFile(Optional sFilename As String)
    Dim lindex As Integer
    Dim ts As TextStream
    
    If sFilename = "" Then
        If mFileName = "" Then
            mnuSaveAs_Click
            Exit Sub
        Else
            sFilename = mFileName
        End If
    End If
    
    With New FileSystemObject
        Set ts = .OpenTextFile(sFilename, ForWriting, True)
    
        For lindex = 1 To UBound(Lines)
            ts.WriteLine Lines(lindex).Text
        Next
        ts.Close
    End With
End Sub

Private Sub AutoSave()
    Dim iUbound As Integer
    Dim ts As TextStream
    Dim iLineIndex As Integer
    
    If mFileName <> "" Then
        With New FileSystemObject
            Set ts = .OpenTextFile(mFileName, ForWriting)
            iUbound = UBound(Lines)
            For iLineIndex = 1 To iUbound
                ts.WriteLine Lines(iLineIndex).Text
            Next
            ts.Close
        End With
    End If
End Sub

Public Sub NewFile()
    mFileName = ""
    Caption = "Untitled - PDL Editor"
    InitialiseConstants
    InitialiseCursor
    InitialiseLines ""
    bInitialised = True
    Form_Resize
End Sub

Public Property Get FileName() As String
    FileName = mFileName
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim linetext As String
    Dim character As String
    Dim capson As Boolean
    Dim Hilited As Boolean
    Dim shifton As Boolean
    Dim ctrlon As Boolean
    Static modifier As String
    Dim bClearModifier As Boolean
    
    capson = (GetKeyState(20) And &H1) = 1
    shifton = (Shift And 1) <> 0
    ctrlon = (Shift And 2) <> 0

    Select Case KeyCode
        Case 16, 17, 18, 19, 20
            Exit Sub
    End Select
    
    If MouseMode <> mdmNotDown Then
        Exit Sub
    End If
    
    If ctrlon Then
        Select Case KeyCode
            Case 65  ' a
                modifier = "a"
            Case 79 ' o
                modifier = "o"
            Case 191 ' /
                modifier = "/"
            Case 223 ' `
                modifier = "`"
            Case 50 ' "
                modifier = """"""
            Case 54 ' ^
                modifier = "^"
            Case 189 ' underscore
                modifier = "_"
            Case 192 ' '
                modifier = "'"
            Case 222 ' ~
                modifier = "~"
            Case 188 ' ,
                modifier = ","
            Case 186
                modifier = ":"
        End Select
        If modifier <> "" Then
            Exit Sub
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyTab
            If HiliteOn Then
                ShiftHiliteText Not shifton
            Else
                InsertText String$(4 - ((Cursor.X - 1) Mod 4), " ")
            End If
            AutoSave
        Case vbKeyEnd
            StartHilite shifton
            Cursor.X = Len(Lines(Cursor.Y).Text) + 1
            ShowHiliteOrCursor
        Case vbKeyHome
            StartHilite shifton
            Cursor.X = 1
            ShowHiliteOrCursor
        Case vbKeyUp
            If Cursor.Y > 1 Then
                StartHilite shifton
                Cursor.Y = Cursor.Y - 1
                If Cursor.Y <= (mTopLineOffset) Then
                    VertScroll.Value = Cursor.Y - 1
                End If
                #If SyntaxCheck <> 0 Then
                    AdjustAll
                #End If
                ShowHiliteOrCursor
            End If
        Case vbKeyDown
            If Cursor.Y < UBound(Lines) Then
                StartHilite shifton
                Cursor.Y = Cursor.Y + 1
                If Cursor.Y >= (mTopLineOffset + VisibleLines - 1) Then
                    VertScroll.Value = Cursor.Y - VisibleLines + 1
                End If
                #If SyntaxCheck <> 0 Then
                    AdjustAll
                #End If
                ShowHiliteOrCursor
            End If
        Case vbKeyLeft
            StartHilite shifton
            If Cursor.X > 1 Then
                Cursor.X = Cursor.X - 1
            ElseIf Cursor.Y > 1 Then
                Cursor.Y = Cursor.Y - 1
                Cursor.X = Len(Lines(Cursor.Y).Text) + 1
                #If SyntaxCheck <> 0 Then
                    AdjustAll
                #End If
            End If
            ShowHiliteOrCursor
        Case vbKeyRight
            StartHilite shifton
            If Cursor.X < (Len(Lines(Cursor.Y).Text) + 1) Then
                Cursor.X = Cursor.X + 1
            ElseIf Cursor.Y < UBound(Lines) Then
                Cursor.Y = Cursor.Y + 1
                Cursor.X = 1
                #If SyntaxCheck <> 0 Then
                    AdjustAll
                #End If
            End If
            ShowHiliteOrCursor

        Case vbKeyReturn
            bTextChanged = True
            If HiliteOn Then
                DeleteHiliteText
            End If
            linetext = Lines(Cursor.Y).Text
            InsertLine Cursor.Y + 1
            Lines(Cursor.Y + 1).Text = Mid$(linetext, Cursor.X)
            Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1)
            RenderLine Cursor.Y
            bTextChanged = True
            #If SyntaxCheck <> 0 Then
                AdjustAll
            #End If
            Cursor = position(1, Cursor.Y + 1)
            PositionCaret Cursor
            AutoSave
        Case vbKeyDelete ' Delete right
            bTextChanged = True
            If HiliteOn Then
                DeleteHiliteText
            Else
                linetext = Lines(Cursor.Y).Text
                If Cursor.X = (Len(linetext) + 1) Then
                    If Cursor.Y < (UBound(Lines) - 1) Then
                        ClearCodes Cursor.Y
                        Lines(Cursor.Y).Text = linetext & Lines(Cursor.Y + 1).Text
                        RemoveLine (Cursor.Y + 1)
                        #If SyntaxCheck <> 0 Then
                            AdjustAll
                        #End If
                        PositionCaret Cursor
                    End If
                Else
                    ClearCodes Cursor.Y
                    If Not ctrlon Then
                        Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1) & Mid$(linetext, Cursor.X + 1)
                    Else
                        Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1) & LTrim$(Mid$(linetext, Cursor.X + 1))
                    End If
                    RenderLine Cursor.Y
                    PositionCaret Cursor
                End If
            End If
            AutoSave
        Case vbKeyBack
            bTextChanged = True
            If HiliteOn Then
                DeleteHiliteText
            Else
                linetext = Lines(Cursor.Y).Text
                If Cursor.X = 1 Then
                    If Cursor.Y > 1 Then
                        Cursor = position(Len(Lines(Cursor.Y - 1).Text) + 1, Cursor.Y - 1)
                        ClearCodes Cursor.Y
                        Lines(Cursor.Y).Text = Lines(Cursor.Y).Text & linetext
                        If Cursor.Y < (UBound(Lines) - 1) Then
                            RemoveLine (Cursor.Y + 1)
                        End If
                        #If SyntaxCheck <> 0 Then
                            AdjustAll
                        #End If
                        PositionCaret Cursor
                    End If
                Else
                    ClearCodes Cursor.Y
                    If Not ctrlon Then
                        Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 2) & Mid$(linetext, Cursor.X)
                    Else
                        Dim tempx As Integer
                        tempx = Cursor.X
                        Cursor.X = Cursor.X - (Len(Left$(linetext, Cursor.X - 2)) - Len(RTrim$(Left$(linetext, Cursor.X - 2))))
                        Lines(Cursor.Y).Text = RTrim$(Left$(linetext, tempx - 2)) & Mid$(linetext, tempx)
                    End If
                    RenderLine Cursor.Y
                    Cursor.X = Cursor.X - 1
                    PositionCaret Cursor
                End If
            End If
            AutoSave
        Case vbKeyA To vbKeyZ, vbKeySpace
            If Shift = 2 And KeyCode = vbKeyC Then
                Clipboard.Clear
                Clipboard.SetText GetHiliteText
            ElseIf Shift = 2 And KeyCode = vbKeyX Then
                Clipboard.Clear
                Clipboard.SetText GetHiliteText
                If HiliteOn Then DeleteHiliteText
                If (Cursor.Y = UBound(Lines)) And (Lines(Cursor.Y).Text = "") Then
                    InsertLine Cursor.Y
                End If
            Else
                If HiliteOn Then DeleteHiliteText
                If (Cursor.Y = UBound(Lines)) And (Lines(Cursor.Y).Text = "") Then
                    InsertLine Cursor.Y
                End If
                Select Case Shift
                    Case 0, 1
                        If modifier = "" Then
                            InsertCharacter Chr$(KeyCode), LCase(Chr$(KeyCode)), 1 - Abs(Shift + capson)
                        Else
                            Dim iAccentIndex As Integer
                            Dim iAccentedLetterIndex As Integer
                            Dim sChar As String
                            
                            iAccentIndex = InStr(vAccentIndex, modifier)
                            iAccentedLetterIndex = InStr(vAccentedLetterIndex, LCase$(Chr$(KeyCode)))
                            
                            If iAccentedLetterIndex = 0 Or iAccentIndex = 0 Then
                                InsertCharacter Chr$(KeyCode), LCase(Chr$(KeyCode)), 1 - Abs(Shift + capson)
                            Else
                                sChar = Mid$(vAccented(iAccentedLetterIndex - 1)(0), iAccentIndex, 1)
                                If sChar <> "x" Then
                                    InsertCharacter sChar, Mid$(vAccented(iAccentedLetterIndex - 1)(1), iAccentIndex, 1), 1 - Abs(Shift + capson)
                                Else
                                    InsertCharacter Chr$(KeyCode), LCase(Chr$(KeyCode)), 1 - Abs(Shift + capson)
                                End If
                            End If
                        End If
                    Case 2
                        Select Case KeyCode
                            Case vbKeyY
                                RemoveLine Cursor.Y
                            Case vbKeyV
                                InsertText Clipboard.GetText
                        End Select
                End Select
                AutoSave
            End If
        Case vbKey0 To vbKey9
            If HiliteOn Then DeleteHiliteText
            If ctrlon And KeyCode = 49 Then
                InsertCharacter "°"
            Else
                InsertCharacter Chr$(KeyCode), Mid$(")!""£$%^&*(", KeyCode - 47, 1), Shift And 1
            End If
        Case 186 To 192
            If HiliteOn Then DeleteHiliteText
            If modifier = "/" And KeyCode = 191 Then
                InsertCharacter "ø"
            Else
                InsertCharacter Mid$(";=,-./'", KeyCode - 185, 1), Mid$(":+<_>?@", KeyCode - 185, 1), Shift And 1
            End If
        Case 219 To 223
            If HiliteOn Then DeleteHiliteText
            InsertCharacter Mid$("[\]#`", KeyCode - 218, 1), Mid$("{|}~", KeyCode - 218, 1), Shift And 1
        Case vbKeyNumpad0 To vbKeyNumpad9
            If HiliteOn Then DeleteHiliteText
            InsertCharacter Mid$("0123456789", KeyCode - vbKeyNumpad0 + 1, 1)
        Case vbKeyAdd
            If HiliteOn Then DeleteHiliteText
            InsertCharacter "+"
        Case vbKeySubtract
            If HiliteOn Then DeleteHiliteText
            InsertCharacter "-"
        Case vbKeyDivide
            If HiliteOn Then DeleteHiliteText
            InsertCharacter "/"
        Case vbKeyMultiply
            If HiliteOn Then DeleteHiliteText
            InsertCharacter "*"
        Case vbKeyDecimal
            If HiliteOn Then DeleteHiliteText
            InsertCharacter "."
    End Select
    modifier = ""
    AdjustScreenForCursor
    ShowCursor
End Sub

Private Sub StartHilite(ByVal shifton As Boolean)
    If shifton Then
        If Not HiliteOn Then
            OldCursor = Cursor
            HiliteStart = Cursor
            HiliteOn = True
        End If
    Else
        RemoveHilite
    End If
End Sub

Private Sub ShowHiliteOrCursor()
    If HiliteOn Then
        DynamicHilite
    Else
        PositionCursor Cursor
    End If
End Sub

Private Sub AdjustScreenForCursor()
    Dim abscursor As AbsoluteCoord
    
    abscursor = AbsolutePos(Cursor)
    
    If Cursor.Y <= (mTopLineOffset) Then
        VertScroll.Value = Cursor.Y - 1
    End If
    
    If Cursor.Y >= (mTopLineOffset + VisibleLines - 1) Then
        VertScroll.Value = Cursor.Y - VisibleLines + 1
    End If
    
    If abscursor.X < 0 Then
        HorizScroll.Value = abscursor.X / Screen.TwipsPerPixelX + HorizScroll.Value
        Me.Refresh
    End If
    
    If (abscursor.X + LeftMargin + 50) > (Me.ScaleWidth - VertScroll.Width) Then
        HorizScroll.Value = (abscursor.X + LeftMargin + 50 + VertScroll.Width - mLeftSide - Me.ScaleWidth) / Screen.TwipsPerPixelX
        Me.Refresh
    End If
End Sub

Private Function StripQuotes(sString) As String
    Dim oTree As New ParseTree
    
    Stream.Text = sString
    If oStripQuotes.Parse(oTree) Then
        StripQuotes = oTree.Text
    End If
End Function

Private Sub Form_DblClick()
    Dim thispos As Coord
    Dim leftx As Long
    Dim rightx As Long
    Dim linetext As String
    Const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_0123456789"
    
    linetext = Lines(Cursor.Y).Text
    
    If linetext = "" Then
        Exit Sub
    End If
    
    leftx = Cursor.X
    rightx = Cursor.X
    Do While InStr(alphabet, Mid$(linetext, leftx, 1)) > 0 And leftx > 0
        leftx = leftx - 1
        If leftx = 0 Then
            Exit Do
        End If
    Loop
    While InStr(alphabet, Mid$(linetext, rightx, 1)) > 0 And rightx <= Len(linetext)
        rightx = rightx + 1
    Wend
    
    If HiliteOn Then
        HiliteText HiliteStart, OldCursor, True
        HiliteOn = False
        HiliteStart = Cursor
    End If
        
    HiliteStart.X = leftx + 1
    HiliteStart.Y = Cursor.Y
    Cursor.X = rightx
    OldCursor.X = HiliteStart.X
    OldCursor.Y = HiliteStart.Y
    HiliteOn = True
    DynamicHilite
    MouseMode = mdmHLWord
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mpos As AbsoluteCoord
    Dim sThisPos As Coord
    
    Select Case Button
        Case vbLeftButton
            Select Case MouseMode
                Case mdmNotDown
                    Select Case CursorMode
                        Case cmNormal
                            MouseMode = mdmHL
                            Form_MouseMove Button, Shift, X, Y
                        Case cmOverHL
                        Case cmMargin
                            MouseMode = mdmHLLine
                            Form_MouseMove Button, Shift, X, Y
                        Case cmDragMove, cmDragCopy
                            MouseMode = mdmHLDragging
                            Form_MouseMove Button, Shift, X, Y
                    End Select
                Case mdmHL, mdmHLWord
                    mpos.X = X - LeftMargin
                    mpos.Y = Y
                    Cursor = TextPos(mpos)
                    If Not HiliteOn Then
                        HiliteOn = True
                    End If
                    DynamicHilite
                Case mdmHLLine
                    mpos.X = X - LeftMargin
                    mpos.Y = Y
                    Cursor = TextPos(mpos)
                    
                    If Not HiliteOn Then
                        HiliteOn = True
                    End If
                    
                    If Cursor.Y < iLineHighlightIndex Then
                        Cursor.X = 1
                    ElseIf Cursor.Y > iLineHighlightIndex Then
                        Cursor.X = Len(Lines(Cursor.Y).Text) + 1
                    Else
                        Cursor.X = Len(Lines(iLineHighlightIndex).Text) + 1
                    End If
                    DynamicHilite
                Case mdmHLDragging
                    Call SetCursorModeDrag((Shift And 2) = 0)
                    mpos.X = X - LeftMargin
                    mpos.Y = Y
                    Cursor = TextPos(mpos)
                    ShowCursor
            End Select
            
        Case vbRightButton
        Case Else ' no button
            If X < LeftMargin Then
                CursorMode = cmMargin
            Else
                If MouseOverHilite(X, Y) Then
                    CursorMode = cmOverHL
                Else
                    CursorMode = cmNormal
                End If
            End If
    End Select
    'DisplayMouseMode
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mpos As AbsoluteCoord
    Dim OldCursorY As Integer

    Select Case Button
        Case vbLeftButton
            Select Case CursorMode
                Case cmOverHL
                    sDragText = GetHiliteText
                    Call SetCursorModeDrag((Shift And 1) <> 0)
                        
                Case cmMargin
                    RemoveHilite
                    mpos.X = X - LeftMargin
                    mpos.Y = Y
                    
                    HiliteStart = TextPos(mpos)
                    HiliteStart.X = 1
                    OldCursor = HiliteStart
                    Cursor = HiliteStart
                    Cursor.X = Len(Lines(HiliteStart.Y).Text) + 1
                    
                    iLineHighlightIndex = HiliteStart.Y
                    HiliteOn = True
                    HiliteText Cursor, HiliteStart
                    MouseMode = mdmHLLine
                    
                Case cmNormal
                    mpos.X = X - LeftMargin
                    mpos.Y = Y
                    OldCursorY = Cursor.Y
                    Cursor = TextPos(mpos)
                    
                    #If SyntaxCheck <> 0 Then
                        If Cursor.Y <> OldCursorY Then
                            If OldCursorY <> 0 Then
                                AdjustAll
                            End If
                        End If
                    #End If
                    
                    ShowCursor
                    RemoveHilite
                    OldCursor = Cursor
                    HiliteStart = Cursor
            End Select
            
        Case vbRightButton
            ' Nothing
        Case Else
            ' Nothing
    End Select
    'DisplayMouseMode
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
            Select Case MouseMode
                Case mdmNotDown
                    ' Nothing
                Case mdmHL, mdmHLWord
                    ' Nothing
                Case mdmHLLine
                    ' Nothing
                Case mdmHLDragging
                    If Not MouseOverHilite(X, Y) Then
                        If LessThan(Cursor, HiliteStart) Then
                            DeleteHiliteText False
                            InsertBlankLine
                            InsertText sDragText
                        Else
                            InsertBlankLine
                            InsertText sDragText
                            DeleteHiliteText False, True
                        End If
                    Else
                        HideCursor
                    End If
            End Select
            MouseMode = mdmNotDown
            Form_MouseMove 0, Shift, X, Y
        Case vbRightButton
        Case Else
    End Select
    'DisplayMouseMode
End Sub

Private Sub DisplayMouseMode()
    Select Case MouseMode
        Case mdmHL
            Caption = "HILITE"
        Case mdmHLWord
            Caption = "HILITEWORD"
        Case mdmHLLine
            Caption = "LINEHILITE"
    End Select
End Sub

Private Sub SetCursorModeDrag(ByVal ctrlon As Boolean)
    If ctrlon Then
        CursorMode = cmDragMove
    Else
        CursorMode = cmDragCopy
    End If
End Sub

Private Sub InsertBlankLine()
    If (Cursor.Y = UBound(Lines)) And (Lines(Cursor.Y).Text = "") Then
        InsertLine Cursor.Y
    End If
End Sub

Private Function MouseOverHilite(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim mpos As AbsoluteCoord
    Dim sThisPos As Coord
    
    If HiliteOn Then
        mpos.X = X - LeftMargin
        mpos.Y = Y
        sThisPos = TextPos(mpos)
        
        If LessThan(OldCursor, HiliteStart) Then
            If LessThan(sThisPos, OldCursor) Then
            ElseIf LessThan(sThisPos, HiliteStart) Then
                MouseOverHilite = True
            End If
        ElseIf LessThan(HiliteStart, OldCursor) Then
            If LessThan(sThisPos, HiliteStart) Then
            ElseIf LessThan(sThisPos, OldCursor) Then
                MouseOverHilite = True
            End If
        End If
    End If
End Function

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Effect
        Case 7
            If Data.GetFormat(15) Then
                LoadFile Data.Files(1)
            End If
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mFileName = "" Then
        mnuSaveAs_Click
    End If
End Sub

Private Sub Form_Paint()
    TopLineOffset = VertScroll.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMenuState
End Sub

Private Sub HorizScroll_Change()
    LeftSide = -HorizScroll.Value * Screen.TwipsPerPixelX
    Me.Line (0, 0)-Step(LeftMargin - 20, Me.ScaleHeight), vbWindowBackground, BF
    PositionCursor Cursor
End Sub

Private Sub mnuFind_Click()
    FindForm.Show
End Sub

Public Sub FindText(sSearchText As String, Optional bSelectedOnly As Boolean = False)
    Dim iLineIndex As Long
    
    For iLineIndex = 1 To UBound(Lines)
        If InStr(Lines(iLineIndex).Text, sSearchText) Then
            'redim preserve lines(ilineindex).Codes
        End If
    Next
End Sub

Private Sub VertScroll_Change()
    TopLineOffset = VertScroll.Value
    PositionCursor Cursor
    Me.Line (0, 0)-Step(LeftMargin, Me.ScaleHeight), vbWindowBackground, BF
End Sub

Private Sub InsertCharacter(ccode As String, Optional acode As String, Optional choice As Integer)
    Dim linetext As String
    
    If choice = 1 Then
        ccode = acode
    End If
    ClearCodes Cursor.Y
    linetext = Lines(Cursor.Y).Text
    Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1) & ccode & Mid$(linetext, Cursor.X)
    bTextChanged = True
    Cursor.X = Cursor.X + 1
    RenderLine Cursor.Y
    PositionCaret Cursor
    AutoSave
End Sub


Private Sub RenderLine(ByVal lineindex As Integer, Optional ByVal startchar As Integer = 1, Optional ByVal endchar As Integer)
    Dim startcharabs As AbsoluteCoord
    Dim endcharabs As AbsoluteCoord
    Dim zero As Long
    Dim previouscharpos As Integer
    Dim charpos As Integer
    Dim codeindex As Integer
    Dim bWholeLine As Boolean
    
    startcharabs = AbsolutePos(position(1, lineindex))
    
    If lineindex > UBound(Lines) Then
        Me.Line (LeftMargin, startcharabs.Y)-Step(Me.ScaleWidth, LineHeight), vbWindowBackground, BF
        Exit Sub
    End If
    
    If endchar = 0 Then
        endchar = Len(Lines(lineindex).Text)
        bWholeLine = True
    End If
    
    If endchar < startchar Then
        Me.Line (LeftMargin, startcharabs.Y)-Step(Me.ScaleWidth, LineHeight), vbWindowBackground, BF
        Exit Sub
    End If
    
    startcharabs = AbsolutePos(position(startchar, lineindex))
    endcharabs = AbsolutePos(position(endchar, lineindex))
    
    If bWholeLine Then
        Me.Line (LeftMargin, startcharabs.Y)-Step(Me.ScaleWidth, LineHeight), vbWindowBackground, BF
    Else
        Me.Line (startcharabs.X + LeftMargin, startcharabs.Y)-Step(endcharabs.X - startcharabs.X, LineHeight), vbWindowBackground, BF
    End If
    Me.Line (startcharabs.X + LeftMargin, startcharabs.Y)-Step(0, 0)
    
    CopyMemory zero, ByVal (VarPtr(Lines(lineindex)) + 4), 4&
    Me.ForeColor = vbWindowText
    
    If zero = 0 Then
        Me.Print Mid$(Lines(lineindex).Text, startchar, endchar - startchar + 1)
        Exit Sub
    End If
    
    previouscharpos = startchar
    
    codeindex = 0
    Do While (codeindex <= UBound(Lines(lineindex).Codes))
        If (Lines(lineindex).Codes(codeindex).position < startchar) Then
            Me.ForeColor = Lines(lineindex).Codes(codeindex).colour
            codeindex = codeindex + 1
        Else
            Exit Do
        End If
    Loop

    charpos = startchar
    While codeindex <= UBound(Lines(lineindex).Codes) And charpos <= endchar
        charpos = Lines(lineindex).Codes(codeindex).position
        If charpos >= endchar Then
            charpos = endchar
        End If
        Me.Print Mid$(Lines(lineindex).Text, previouscharpos, charpos - previouscharpos);
        Me.ForeColor = Lines(lineindex).Codes(codeindex).colour
        previouscharpos = charpos
        codeindex = codeindex + 1
    Wend
    Me.Print Mid$(Lines(lineindex).Text, charpos, endchar - charpos + 1)
End Sub

#If SyntaxCheck <> 0 Then
Private Sub AdjustAll()
    Dim sProgram As String
    Dim iLineIndex As Integer
    Dim iNext As Integer
    Dim vAdjustedLine As Variant
    Dim iPositionIndex As Integer
    Dim oPositions As Collection
    Dim oColours As Collection
    Dim iLineOffset As Integer
    Dim vLine As Variant
    Dim iLineLen As Integer
    Dim iPositionValue As Integer
    Dim lCurrentColour As Long
    Dim iParseLength As Integer
    Dim bParsedOk As Boolean
    Dim bOK As Boolean
    
    If Not bTextChanged Then
        Exit Sub
    End If
    bTextChanged = False
    
    Me.MousePointer = vbHourglass
    
    For iLineIndex = 1 To UBound(Lines)
        sProgram = sProgram & Lines(iLineIndex).Text & vbCrLf
    Next
    
    ReDim Lines(1 To 1) As LineCode
    
    iNext = 1
    iLineIndex = 1
    Do
        vAdjustedLine = ParseLine(Mid$(sProgram, iNext), iParseLength, bParsedOk)
        If bParsedOk Then
            iNext = iNext + iParseLength - 1
        Else
            If InStr(vAdjustedLine(0), vbCrLf) Then
                iNext = iNext + InStr(vAdjustedLine(0), vbCrLf) + 1
                vAdjustedLine(0) = Left$(vAdjustedLine(0), InStr(vAdjustedLine(0), vbCrLf) - 1)
            Else
                iNext = iNext + Len(vAdjustedLine(0))
            End If
        End If
        
        Set oPositions = vAdjustedLine(1)
        Set oColours = vAdjustedLine(2)
        iPositionIndex = 1
        
        iLineOffset = 1
        If vAdjustedLine(0) = "" Then
            ReDim Preserve Lines(1 To iLineIndex) As LineCode
            iLineIndex = iLineIndex + 1
        End If
        
        For Each vLine In Split(vAdjustedLine(0), vbCrLf)
            iLineLen = Len(vLine)
            
            ReDim Preserve Lines(1 To iLineIndex) As LineCode
            Lines(iLineIndex).Text = vLine
            ReDim Lines(iLineIndex).Codes(0) As ColourInfo
            Lines(iLineIndex).CodesNotEmpty = True
            
            bOK = True
            While bOK
                If iPositionIndex <= oPositions.Count Then
                    iPositionValue = oPositions(iPositionIndex)
                    If iPositionValue < iLineOffset Then
                        iPositionIndex = iPositionIndex + 1
                    Else
                        If iPositionValue >= iLineOffset And iPositionValue <= (iLineOffset + iLineLen - 1) Then
                            ReDim Preserve Lines(iLineIndex).Codes(UBound(Lines(iLineIndex).Codes) + 1) As ColourInfo
                            With Lines(iLineIndex).Codes(UBound(Lines(iLineIndex).Codes))
                                .colour = oColours(iPositionIndex)
                                .position = iPositionValue - iLineOffset + 1
                                lCurrentColour = .colour
                            End With
                            iPositionIndex = iPositionIndex + 1
                        Else
                            bOK = False
                        End If
                    End If
                Else
                    bOK = False
                End If
            Wend
            
            RenderLine iLineIndex
            iLineIndex = iLineIndex + 1
            iLineOffset = iLineOffset + iLineLen + 2
        Next
    Loop Until iNext > Len(sProgram)
    AutoSave
    Me.MousePointer = vbIbeam
End Sub
#End If

Private Sub ClearCodes(ByVal lineindex As Long)
    #If SyntaxCheck <> 0 Then
        Dim zero As Long
        
        CopyMemory zero, ByVal (VarPtr(Lines(lineindex)) + 4), 4&
        If zero <> 0 Then
            Erase Lines(lineindex).Codes
            Lines(lineindex).CodesNotEmpty = False
            RenderLine lineindex
        End If
    #End If
End Sub

Private Sub InsertText(ByVal insert As String)
    Dim splitline As Variant
    Dim maxlineinsert As Long
    Dim lineinsert As Long
    Dim linetext As String
    Dim tail As String
    
    splitline = Split(insert, vbCrLf)
    maxlineinsert = UBound(splitline)

    linetext = Lines(Cursor.Y).Text
    tail = Mid$(linetext, Cursor.X)
    For lineinsert = 0 To UBound(splitline)
        If lineinsert = 0 And lineinsert = UBound(splitline) Then
            ClearCodes Cursor.Y
            linetext = Lines(Cursor.Y).Text
            Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1) & splitline(lineinsert) & tail
            Cursor.X = Cursor.X + Len(splitline(lineinsert))
        ElseIf lineinsert = 0 Then
            ClearCodes Cursor.Y
            linetext = Lines(Cursor.Y).Text
            Lines(Cursor.Y).Text = Left$(linetext, Cursor.X - 1) & splitline(lineinsert)
        ElseIf lineinsert = UBound(splitline) Then
            Cursor.X = 1
            Cursor.Y = Cursor.Y + 1
            ClearCodes Cursor.Y
            InsertLine Cursor.Y
            linetext = Lines(Cursor.Y).Text
            Lines(Cursor.Y).Text = splitline(lineinsert) & tail
            Cursor.X = Len(splitline(lineinsert)) + 1
        Else
            Cursor.X = 1
            Cursor.Y = Cursor.Y + 1
            ClearCodes Cursor.Y
            InsertLine Cursor.Y
            Lines(Cursor.Y).Text = splitline(lineinsert)
        End If
        
        RenderLine Cursor.Y
    Next

    PositionCaret Cursor
    bTextChanged = True
    AutoSave
End Sub


Private Sub InsertLine(ByVal lineindex As Integer)
    Dim index As Integer

    ReDim Preserve Lines(1 To UBound(Lines) + 1) As LineCode
    
    For index = UBound(Lines) To lineindex + 1 Step -1
        Lines(index) = Lines(index - 1)
        RenderLine index
    Next
    Lines(lineindex).Text = ""
    RenderLine lineindex
    AutoSave
End Sub


Private Sub RemoveLine(ByVal lineindex As Integer)
    Dim index As Long
    For index = lineindex To UBound(Lines) - 1
        Lines(index) = Lines(index + 1)
        RenderLine index
    Next
    ReDim Preserve Lines(1 To UBound(Lines) - 1) As LineCode
    AutoSave
End Sub

Private Sub HiliteText(HLStartIn As Coord, HLEndIn As Coord, Optional ByVal bRemove As Boolean)
    Dim linecount As Integer
    Dim absstart As AbsoluteCoord
    Dim absend As AbsoluteCoord
    Dim startpos As Integer
    Dim endpos As Integer
    Dim HLStart As Coord
    Dim HLEnd As Coord
    
    If LessThan(HLEndIn, HLStartIn) Then
        HLStart = HLEndIn
        HLEnd = HLStartIn
    Else
        HLStart = HLStartIn
        HLEnd = HLEndIn
    End If
    
    For linecount = HLStart.Y To HLEnd.Y
        absstart = AbsolutePos(position(1, linecount)): startpos = 1
        absend = AbsolutePos(position(Len(Lines(linecount).Text) + 1, linecount)): endpos = Len(Lines(linecount).Text) + 1
        
        If linecount = HLStart.Y Then absstart = AbsolutePos(HLStart): startpos = HLStart.X
        If linecount = HLEnd.Y Then absend = AbsolutePos(HLEnd): endpos = HLEnd.X
        
        If Not bRemove Then
            Me.ForeColor = vbHighlightText
            Me.Line (LeftMargin + absstart.X + mLeftSide, absstart.Y)-Step(absend.X - absstart.X, LineHeight), vbHighlight, BF
            Me.Line (LeftMargin + mLeftSide + absstart.X, absstart.Y)-Step(0, 0)
            Me.Print Mid$(Lines(linecount).Text, startpos, endpos - startpos)
        Else
            RenderLine linecount, startpos, endpos
        End If
    Next
End Sub

Private Sub DynamicHilite()
    Dim bShowCaret As Boolean
    
    If LessThan(OldCursor, HiliteStart) Then
        If LessThan(HiliteStart, Cursor) Then
            HiliteText OldCursor, HiliteStart, True
            HiliteText HiliteStart, Cursor
        ElseIf LessThan(Cursor, HiliteStart) Then
            If LessThan(Cursor, OldCursor) Then
                HiliteText Cursor, OldCursor
            ElseIf LessThan(OldCursor, Cursor) Then
                HiliteText OldCursor, Cursor, True
            End If
        Else
            HiliteText OldCursor, Cursor, True
        End If
    ElseIf LessThan(HiliteStart, OldCursor) Then
        If LessThan(HiliteStart, Cursor) Then
            If LessThan(Cursor, OldCursor) Then
                HiliteText Cursor, OldCursor, True
            ElseIf LessThan(OldCursor, Cursor) Then
                HiliteText OldCursor, Cursor
            End If
        ElseIf LessThan(Cursor, HiliteStart) Then
            HiliteText HiliteStart, OldCursor, True
            HiliteText Cursor, HiliteStart
        Else
            HiliteText Cursor, OldCursor, True
        End If
    Else
        If LessThan(Cursor, HiliteStart) Then
            HiliteText Cursor, HiliteStart
        ElseIf LessThan(HiliteStart, Cursor) Then
            HiliteText HiliteStart, Cursor
        Else
            bShowCaret = True
        End If
    End If
    
    If bShowCaret Then
        If Not bCaretVisible Then
            ShowCaret Me.hwnd
            bCaretVisible = True
        End If
    Else
        If bCaretVisible Then
            HideCaret Me.hwnd
            bCaretVisible = False
        End If
    End If
    
    OldCursor = Cursor
End Sub

Private Sub HideCursor()
    If bCaretVisible Then
        HideCaret Me.hwnd
        bCaretVisible = False
    End If
End Sub

Private Function GetHiliteText() As String
    Dim lineindex As Long
    Dim linetext As String
    Dim ThisHiliteStart As Coord
    Dim ThisHiliteEnd As Coord
    
    If LessThan(HiliteStart, Cursor) Then
        ThisHiliteStart = HiliteStart
        ThisHiliteEnd = Cursor
    ElseIf Equal(HiliteStart, Cursor) Then
        Exit Function
    Else
        ThisHiliteStart = Cursor
        ThisHiliteEnd = HiliteStart
    End If
    
    For lineindex = ThisHiliteStart.Y To ThisHiliteEnd.Y
        linetext = Lines(lineindex).Text
        If lineindex = ThisHiliteStart.Y And lineindex = ThisHiliteEnd.Y Then
            GetHiliteText = Mid$(linetext, ThisHiliteStart.X, ThisHiliteEnd.X - ThisHiliteStart.X)
        ElseIf lineindex = ThisHiliteStart.Y Then
            GetHiliteText = Mid$(linetext, ThisHiliteStart.X) & vbCrLf
        ElseIf lineindex = ThisHiliteEnd.Y Then
            GetHiliteText = GetHiliteText & Left$(linetext, ThisHiliteEnd.X - 1)
        Else
            GetHiliteText = GetHiliteText & linetext & vbCrLf
        End If
    Next
End Function

Private Sub ShiftHiliteText(bRight As Boolean)
    Dim ThisHiliteStart As Coord
    Dim ThisHiliteEnd As Coord
    Dim lLineIndex As Long
    Dim iCodeIndex As Integer
    Dim iSpaces As Integer
    
    If Not HiliteOn Then
        Exit Sub
    End If
    
    If LessThan(HiliteStart, Cursor) Then
        ThisHiliteStart = HiliteStart
        ThisHiliteEnd = Cursor
    ElseIf Equal(HiliteStart, Cursor) Then
        Exit Sub
    Else
        ThisHiliteStart = Cursor
        ThisHiliteEnd = HiliteStart
    End If
    
    If ThisHiliteStart.Y = 0 Then
        Exit Sub
    End If
    
    If bRight Then
        For lLineIndex = ThisHiliteStart.Y To ThisHiliteEnd.Y
            Lines(lLineIndex).Text = "    " & Lines(lLineIndex).Text
            If Lines(lLineIndex).CodesNotEmpty Then
                For iCodeIndex = 1 To UBound(Lines(lLineIndex).Codes)
                    Lines(lLineIndex).Codes(iCodeIndex).position = Lines(lLineIndex).Codes(iCodeIndex).position + 4
                Next
            End If
        Next
        ThisHiliteStart.X = 1
        ThisHiliteEnd.X = Len(Lines(ThisHiliteEnd.Y).Text)
    Else
        For lLineIndex = ThisHiliteStart.Y To ThisHiliteEnd.Y
            iSpaces = CountSpacesOnLeft(Lines(lLineIndex).Text)
            Lines(lLineIndex).Text = Mid$(Lines(lLineIndex).Text, iSpaces + 1)
            If Lines(lLineIndex).CodesNotEmpty Then
                For iCodeIndex = 1 To UBound(Lines(lLineIndex).Codes)
                    Lines(lLineIndex).Codes(iCodeIndex).position = Lines(lLineIndex).Codes(iCodeIndex).position - iSpaces
                Next
            End If
            RenderLine lLineIndex
        Next
        ThisHiliteStart.X = 1
        ThisHiliteEnd.X = Len(Lines(ThisHiliteEnd.Y).Text)
    End If
    
    HiliteText ThisHiliteStart, ThisHiliteEnd
    HiliteStart = ThisHiliteStart
    Cursor = ThisHiliteEnd
    OldCursor = Cursor
End Sub

Private Function CountSpacesOnLeft(sText As String)
    CountSpacesOnLeft = Len(sText) - Len(LTrim$(sText))
    If CountSpacesOnLeft > 4 Then
        CountSpacesOnLeft = 4
    End If
End Function

Private Function EmptyArray(ByRef vArray As Variant) As Boolean
    Dim zero As Long
    CopyMemory zero, ByVal (VarPtr(vArray) + 4), 4&
    EmptyArray = zero = 0
End Function

Private Sub DeleteHiliteText(Optional ByVal bRepositionCursor As Boolean = True, Optional ByVal bMoveCursor As Boolean = False)
    Dim ThisHiliteStart As Coord
    Dim ThisHiliteEnd As Coord
    Dim lineindex As Integer
    Dim startpos As Integer
    Dim endpos As Integer
    Dim linetext As String
    Dim linetext1 As String
    Dim linetext2 As String
    
    If Not HiliteOn Then
        Exit Sub
    End If
    
    HiliteOn = False
    
    bTextChanged = True
    
    If LessThan(HiliteStart, OldCursor) Then
        ThisHiliteStart = HiliteStart
        ThisHiliteEnd = OldCursor
    ElseIf Equal(HiliteStart, OldCursor) Then
        Exit Sub
    Else
        ThisHiliteStart = OldCursor
        ThisHiliteEnd = HiliteStart
    End If
    
    If ThisHiliteStart.Y = 0 Then
        Exit Sub
    End If
    
    If ThisHiliteStart.Y = ThisHiliteEnd.Y Then
        startpos = ThisHiliteStart.X
        endpos = ThisHiliteEnd.X
        linetext = Lines(ThisHiliteStart.Y).Text
        Lines(ThisHiliteStart.Y).Text = Left$(linetext, startpos - 1) & Mid$(linetext, endpos)
        RenderLine ThisHiliteStart.Y
        If bMoveCursor Then
            Cursor.X = Cursor.X - (endpos - startpos)
        End If
    Else
        startpos = ThisHiliteStart.X
        endpos = ThisHiliteEnd.X
        linetext1 = Lines(ThisHiliteStart.Y).Text
        linetext2 = Lines(ThisHiliteEnd.Y).Text
        Lines(ThisHiliteStart.Y).Text = Left$(linetext1, startpos - 1) & Mid$(linetext2, endpos)
        RenderLine ThisHiliteStart.Y
        
        Dim totallines As Integer
        totallines = ThisHiliteEnd.Y - ThisHiliteStart.Y
        If totallines > 0 Then
            For lineindex = 1 To totallines
                RemoveLine ThisHiliteStart.Y + 1
            Next
        End If
        #If SyntaxCheck <> 0 Then
            AdjustAll
        #End If
        
        If bMoveCursor Then
            Cursor.X = Cursor.X - (endpos - startpos)
            Cursor.Y = Cursor.Y - (ThisHiliteEnd.Y - ThisHiliteStart.Y)
        End If
    End If

    If bRepositionCursor Then
        Cursor = ThisHiliteStart
    End If
    
    If bMoveCursor Or bRepositionCursor Then
        PositionCaret Cursor
    End If
    AutoSave
End Sub

Private Sub RemoveHilite()
    If HiliteOn Then
        HiliteText HiliteStart, OldCursor, True
        HiliteOn = False
        HiliteStart = Cursor
    End If
End Sub

Private Sub PositionCursor(newcursor As Coord)
    Cursor = TextPos(AbsolutePos(newcursor))
    PositionCaret Cursor
End Sub

Private Sub PositionCaret(pos As Coord)
    Dim ap As AbsoluteCoord
    ap = AbsolutePos(pos)
    SetCaretPos (LeftMargin + ap.X) / Screen.TwipsPerPixelX, ap.Y / Screen.TwipsPerPixelY
End Sub

Private Sub ShowCursor()
    If Not bCaretVisible Then
        ShowCaret Me.hwnd
        bCaretVisible = True
    End If
    PositionCaret Cursor
End Sub

Private Function TextPos(pos As AbsoluteCoord) As Coord
    Dim linewidth As Single
    Dim xsearch As Integer
    Dim charwidth As Single
    
    TextPos.Y = Int(pos.Y / LineHeight) + 1 + mTopLineOffset
    
    If TextPos.Y <= UBound(Lines) And TextPos.Y >= LBound(Lines) Then
        TextPos.X = 1
        For xsearch = 1 To Len(Lines(TextPos.Y).Text)
            charwidth = Me.TextWidth(Mid$(Lines(TextPos.Y).Text, xsearch, 1))
            linewidth = linewidth + charwidth
            If linewidth <= (pos.X - mLeftSide + charwidth / 2) Then
                TextPos.X = xsearch + 1
            End If
        Next
    Else
        TextPos.Y = UBound(Lines)
        TextPos.X = 1
    End If
End Function

Private Function AbsolutePos(pos As Coord) As AbsoluteCoord
    Dim linewidth As Single
    Dim xsearch As Integer
    
    AbsolutePos.Y = (pos.Y - 1 - mTopLineOffset) * LineHeight
    AbsolutePos.X = mLeftSide
    
    If pos.Y <= UBound(Lines) And pos.Y >= LBound(Lines) Then
        AbsolutePos.X = Me.TextWidth(Left$(Lines(pos.Y).Text, pos.X - 1)) + mLeftSide
    End If
End Function

Private Function position(ByVal xpos As Integer, ByVal ypos As Integer) As Coord
    position.X = xpos
    position.Y = ypos
End Function

Private Function Equal(c1 As Coord, c2 As Coord) As Boolean
    Equal = (c1.X = c2.X) And (c1.Y = c2.Y)
End Function

Private Function LessThan(c1 As Coord, c2 As Coord) As Boolean
    If c1.Y < c2.Y Then
        LessThan = True
    ElseIf c1.Y > c2.Y Then
        LessThan = False
    Else
        If c1.X < c2.X Then
            LessThan = True
        Else
            LessThan = False
        End If
    End If
End Function

Private Property Let TopLineOffset(iNewTopLine As Integer)
    Dim linecount As Integer
    
    mTopLineOffset = iNewTopLine
    For linecount = mTopLineOffset + 1 To VisibleLines + mTopLineOffset
        RenderLine linecount
    Next
    Me.Refresh
End Property

Private Property Let LeftSide(iNewLeftSide As Single)
    Dim linecount As Integer
    
    mLeftSide = iNewLeftSide
    For linecount = mTopLineOffset + 1 To VisibleLines + mTopLineOffset
        RenderLine linecount
    Next
End Property

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExport_Click()
    FilePicker.ShowSave
    ExportToVisualBasic FilePicker.FileName, mFileName
End Sub

Private Sub mnuNew_Click()
    mnuSaveAs.Enabled = True
    mnuExport.Enabled = True
    NewFile
End Sub

Private Sub mnuOpen_Click()
    FilePicker.ShowOpen
    If FilePicker.FileName <> "" Then
        mnuSaveAs.Enabled = True
        mnuExport.Enabled = True

        LoadFile FilePicker.FileName
    End If
End Sub

Private Sub mnuSaveAs_Click()
    Dim sFilename As String
    
    FilePicker.FileName = FileName
    FilePicker.ShowSave
    If (FilePicker.Flags And 1024) = 1024 Then
        If FilePicker.FileName = "" Then
            mnuSaveAs_Click
        Else
            sFilename = FilePicker.FileName
            mFileName = Left$(sFilename, InStrRev(sFilename, "\")) & NormaliseFileName(FilePicker.FileTitle)
            SaveFile mFileName
        End If
    End If
End Sub

Private Function NormaliseFileName(ByVal sFileTitle As String) As String
    Dim extension As String
    Dim dotpos As Long
    
    dotpos = InStrRev(sFileTitle, ".")
    
    If dotpos = 0 Then
        NormaliseFileName = sFileTitle & ".pdl"
    Else
        NormaliseFileName = sFileTitle
    End If
End Function

' Set Save/Save As menus
Public Sub SetMenuState()
    If Forms.Count = 2 Then
        mnuSaveAs.Enabled = False
        mnuExport.Enabled = False
    Else
        mnuSaveAs.Enabled = True
        mnuExport.Enabled = True
    End If
End Sub

' UNDO/REDO
Private Sub WriteKeyLog(sKey As String)
    ReDim Preserve KeyLog(KeyLogCount) As String * 1
    KeyLog(KeyLogCount) = sKey
    KeyLogCount = KeyLogCount + 1
End Sub

Private Function ReadKeyLog() As String
    KeyLogCount = KeyLogCount - 1
    ReadKeyLog = KeyLog(KeyLogCount)
    ReDim Preserve KeyLog(KeyLogCount) As String * 1
End Function

