VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Quick Grid"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7080
      Top             =   120
   End
   Begin VB.HScrollBar scrHorizontal 
      Height          =   255
      Left            =   0
      Max             =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   8055
   End
   Begin VB.VScrollBar scrVertical 
      Height          =   2415
      Left            =   8040
      Max             =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox goCanvas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.PictureBox goPanel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   15735
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   15735
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   855
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   8400
      End
      Begin VB.TextBox txtReplace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   855
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   8415
      End
      Begin VB.Label lblFind 
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   495
      End
      Begin VB.Label lblReplace 
         BackStyle       =   0  'Transparent
         Caption         =   "Replace:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   735
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu mnuInsertColumn 
         Caption         =   "Column"
      End
      Begin VB.Menu mnuInsertColumnRight 
         Caption         =   "Column Right"
      End
      Begin VB.Menu mnuInsertRow 
         Caption         =   "Row"
      End
      Begin VB.Menu mnuInsertRowDown 
         Caption         =   "Row Down"
      End
      Begin VB.Menu mnuInsertCellRight 
         Caption         =   "Cell Right"
      End
      Begin VB.Menu mnuInsertCellDown 
         Caption         =   "Cell Down"
      End
      Begin VB.Menu mnuInsertPasteRight 
         Caption         =   "Paste Right"
      End
      Begin VB.Menu mnuInsertPasteDown 
         Caption         =   "Paste Down"
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "Selection"
      Begin VB.Menu mnuSelectionClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSelectionUpperCase 
         Caption         =   "Upper Case"
      End
      Begin VB.Menu mnuSelectionLowerCase 
         Caption         =   "Lower Case"
      End
      Begin VB.Menu mnuSelectionCamelCase 
         Caption         =   "Camel Case"
      End
      Begin VB.Menu mnuSelectionUnderscored 
         Caption         =   "Underscored"
      End
      Begin VB.Menu mnuSelectionSpaced 
         Caption         =   "Spaced"
      End
      Begin VB.Menu mnuSelectionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionFill 
         Caption         =   "Fill"
      End
      Begin VB.Menu mnuSelectionPad 
         Caption         =   "Pad"
      End
      Begin VB.Menu mnuSelectionCommaList 
         Caption         =   "Comma List"
      End
      Begin VB.Menu mnuSelectionCommaListQuotes 
         Caption         =   "Comma List && Quotes"
      End
      Begin VB.Menu mnuSelectionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuSelectionDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSelectionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuSelectionRight 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuSelectionUp 
         Caption         =   "Up"
      End
      Begin VB.Menu mnuSelectionDown 
         Caption         =   "Down"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Begin VB.Menu mnuSelectNone 
         Caption         =   "None"
      End
      Begin VB.Menu mnuSelectCell 
         Caption         =   "Cell"
      End
      Begin VB.Menu mnuSelectRow 
         Caption         =   "Row"
      End
      Begin VB.Menu mnuSelectColumn 
         Caption         =   "Column"
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "Table"
      Begin VB.Menu mnuTableRemoveSpaces 
         Caption         =   "Remove Spaces"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTableRemoveEndSpaces 
         Caption         =   "Remove End Spaces"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTableMergeColumns 
         Caption         =   "Merge Columns"
      End
      Begin VB.Menu mnuTableRemoveBlankLines 
         Caption         =   "Remove Blank Lines"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTableReduceSpaces 
         Caption         =   "Reduce Spaces"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTableReduceTabs 
         Caption         =   "Reduce Tabs"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHTMLTable 
         Caption         =   "HTML Table"
      End
      Begin VB.Menu mnuHTMLFloated 
         Caption         =   "HTML Floated"
      End
      Begin VB.Menu mnuTableHTMLList 
         Caption         =   "HTML List"
      End
      Begin VB.Menu mnuTableSort 
         Caption         =   "Sort"
      End
      Begin VB.Menu mnuTableSortDescending 
         Caption         =   "Sort Descending"
      End
      Begin VB.Menu mnuTableSortNumeric 
         Caption         =   "Sort Numeric"
      End
      Begin VB.Menu mnuTableSortNumericDescending 
         Caption         =   "Sort Numeric Descending"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, ByVal lpwTransKey As String, ByVal fuState As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Enum HoverAreas
    haCell
    haColumnDivider
    haColumnSelector
    haRowSelector
    haTableSelector
    haOther
    haNone
End Enum

Private Type Selection
    ColumnIndex As Long
    RowIndex As Long
    CellCharOffset As Long
    TextCharOffset As Long
    Length As Long
    PixelOffset As Long
End Type


Private Enum ModifyFunctions
    mfUpperCase
    mfLowerCase
    mfCamelCase
    mfUnderscored
    mfSpaced
    mfClear
    mfCopy
End Enum

Private Enum MoveDirections
    mdUp
    mdRight
    mdDown
    mdLeft
End Enum

Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5

Private Const MAPVK_VK_TO_VSC = 0
Private Const MAPVK_VSC_TO_VK = 1
Private Const MAPVK_VK_TO_CHAR = 2
Private Const MAPVK_VSC_TO_VK_EX = 3

Private moDragPointer As StdPicture
Private moDragCopyPointer As StdPicture
Private moSelectCellPointer As StdPicture
Private moDownPointer As StdPicture
Private moRightPointer As StdPicture
Private moDividerPointer As StdPicture
Private moHandOpenPointer As StdPicture
Private moHandClosedPointer As StdPicture

Private moAreaSelection As New AreaSelection
Private moCursor As New Cursor
Private moTableInfo As New TableInfo
Private moRenderer As New Renderer
Private moTheme As New Theme
Private moTable As New Table
Private moCellPosition As New CellPosition
Private moStartCellPosition As New CellPosition
Private moDevice As New Device
Private moChangeHistory As New ChangeHistory
Private moTabStops As New TabStops
Private moParsing As New Parsing
Private moSelectionn As New Selectionn

Private mhaDragSelection As HoverAreas
Private mlPreviousX As Long
Private mlPreviousY As Long
Private mlDragStopIndex As Long
Private moCellSelectionOrigin As CellPosition

Private Type MouseUpValues
    Button As Integer
    X As Single
    Y As Single
    UpDown As Long
End Type

Private mmuvMouse As MouseUpValues

Private Enum CursorStates
    csNone
    csDragText
End Enum

Private mcsCursorState As CursorStates

Private Sub Form_Load()
    Dim vTest As Variant
    
    Initialise
    moRenderer.RenderTable
End Sub

Private Sub Initialise()
    With moDevice
        .hDC(0) = goCanvas.hDC
        .hWnd(0) = goCanvas.hWnd
        .LeftTableOffset(0) = 0
        .TopTableOffset(0) = 0
        Set .moDevice = goCanvas
    End With

    With moTheme
        .mlCellColour = RGB(&HE8, &HF3, &HFF)
        .mlEmptyCellColour = RGB(&HF3, &HF5, &HFF) ' #F5FAFF
        .mlSelectedCellColour = RGB(&HFF, &HBD, &H66) ' #FFBD66
        .mlSelectedText = RGB(&H31, &H61, &HC5) '#316AC5
    End With
    
    moParsing.InitialiseParser
        
    With moTable
        Set .moParsing = moParsing
        .msColumnDelimiter = vbTab
        .msRowDelimiter = vbCrLf
        .Table = Array(Array(""))
        '.Text = "a" & vbTab & "b" & vbTab & "c" & vbTab & "d" & vbCrLf & "e" & vbTab & "f" & vbTab & "g" & vbTab & "h" & vbCrLf & "i" & vbTab & "j" & vbTab & "k" & vbTab & "l" & vbCrLf & "m" & vbTab & "n" & vbTab & "o" & vbTab & "p"
    End With
    
    With moTableInfo
        Set .moTable = moTable
        .mlCellHeight = 20
        .mlCellOffsetLeft = 30
        .mlCellOffsetTop = 20
        .mlCellSeparator = 2
    End With
    
    With moTabStops
        .mlMinimumTabWidth = 20
        .DefaultWidth = 30
    End With
    
    With moCellPosition
        Set .moTableInfo = moTableInfo
        Set .moDevice = moDevice
    End With
    
    With moAreaSelection
        Set .moTableInfo = moTableInfo
        Set .moCellPosition.moTableInfo = .moTableInfo
        Set .moCellPosition.moDevice = moDevice
        Set .moTabStops = moTabStops
    End With
    
    With moCursor
        Set .moDevice = moDevice
        Set .moPosition = moCellPosition
        Set .moTableInfo = moTableInfo
        Set .moTabStops = moTabStops
        Set .moVerticalScroll = scrVertical
        .Initialise
    End With
    
    With moRenderer
        Set .moDevice = moDevice
        Set .moTableInfo = moTableInfo
        Set .Theme = moTheme
        Set .moSelectionn = moSelectionn
        Set .moTabStops = moTabStops
        Set .moVerticalScroll = scrVertical
        Set .moHorizontalScroll = scrHorizontal
        .Initialise
    End With
    
    With moSelectionn
        Set .moTable = moTable
        Set .moCursor = moCursor
    End With
    
    Set moDragPointer = LoadPicture(App.Path & "\DRAGMOVE.CUR")
    Set moDragCopyPointer = LoadPicture(App.Path & "\DRAGCOPY.CUR")
    Set moSelectCellPointer = LoadPicture(App.Path & "\cell_select.cur")
    Set moDownPointer = LoadPicture(App.Path & "\down_arrow.cur")
    Set moRightPointer = LoadPicture(App.Path & "\right_arrow.cur")
    Set moDividerPointer = LoadPicture(App.Path & "\divider.cur")
    Set moHandOpenPointer = LoadPicture(App.Path & "\hand_open.cur")
    Set moHandClosedPointer = LoadPicture(App.Path & "\hand_closed.cur")
    
    mhaDragSelection = haNone
    
    Me.Width = GetSetting("QuickPad", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("QuickPad", "Dimensions", "Height", Me.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "QuickPad", "Dimensions", "Width", Me.Width
    SaveSetting "QuickPad", "Dimensions", "height", Me.Height
End Sub

Private Sub goCanvas_DblClick()
    Dim bLeftShift As Boolean
    Dim bRightShift As Boolean
    Dim bLeftControl As Boolean
    Dim bRightControl As Boolean
    Dim bLeftAlt As Boolean
    Dim bRightAlt As Boolean
    Dim lPosition As Long
    Dim lIndex As Long
    
    bLeftShift = GetKeyState(VK_LSHIFT) < 0
    bRightShift = GetKeyState(VK_RSHIFT) < 0
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    bRightControl = GetKeyState(VK_RCONTROL) < 0
    bLeftAlt = GetKeyState(VK_LMENU) < 0
    bRightAlt = GetKeyState(VK_RMENU) < 0
    
'    If mmuvMouse.UpDown = 1 Then
'        Select Case mmuvMouse.Button
'            Case vbLeftButton
'                moAreaSelection.GetSelection mmuvMouse.X, mmuvMouse.Y
'                Select Case moAreaSelection.mhaAreaSelection
'                    Case haCell
'                        Select Case moSelectionn.SelectionType
'                            Case stText, stNone
'                                moSelectionn.SelectionType = stText
'
'                                lPosition = moAreaSelection.moCellPosition.TextPosition
'
'                                For lIndex = lPosition To 0 Step -1
'                                Next
'                                moSelectionn.StartPosition = lIndex + 1
'                                For lIndex = lPosition To Len(moTable.Text)
'                                Next
'                                moSelectionn.EndPosition = lIndex - 1
'
'                                If moSelectionn.StartPosition = moSelectionn.EndPosition Then
'                                    moSelectionn.SelectionType = stNone
'                                Else
'                                    moSelectionn.SelectionType = stText
'                                End If
'                                moCellPosition.TextPosition = lIndex - 1
'                                Rerender False
'                        End Select
'                End Select
'        End Select
'    End If
End Sub

Private Sub goCanvas_GotFocus()
    moCursor.RecreateCursor
End Sub

Private Sub goCanvas_LostFocus()
    moCursor.HideCursor
    txtReplace.TabStop = True
    goCanvas.TabStop = True
End Sub

Private Sub goCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bLeftShift As Boolean
    Dim bRightShift As Boolean
    Dim bLeftControl As Boolean
    Dim bRightControl As Boolean
    Dim bLeftAlt As Boolean
    Dim bRightAlt As Boolean
    Dim bShift As Boolean
    Dim bControl As Boolean
    Dim lTextPosition As Long
    
    bLeftShift = GetKeyState(VK_LSHIFT) < 0
    bRightShift = GetKeyState(VK_RSHIFT) < 0
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    bRightControl = GetKeyState(VK_RCONTROL) < 0
    bLeftAlt = GetKeyState(VK_LMENU) < 0
    bRightAlt = GetKeyState(VK_RMENU) < 0
    bShift = bLeftShift Or bRightShift
    bControl = bLeftControl Or bRightControl
    
    mmuvMouse.Button = Button
    mmuvMouse.X = X
    mmuvMouse.Y = Y
    mmuvMouse.UpDown = 0
    
    Select Case Button
        Case vbLeftButton
            moAreaSelection.GetSelection X, Y
            
            Select Case moAreaSelection.mhaAreaSelection
                Case haColumnSelector
                    moSelectionn.SelectionType = stColumns
                Case haRowSelector
                    moSelectionn.SelectionType = stRows
                Case haCell
                    If bLeftControl Or bRightControl Then
                        moSelectionn.SelectionType = stCells
                    Else
                        moSelectionn.SelectionType = stText
                    End If
                Case haTableSelector
                    moSelectionn.SelectionType = stTable
            End Select
    
            Select Case moAreaSelection.mhaAreaSelection
                Case haColumnSelector, haRowSelector, haCell
                    If bShift Then
                        If moCellSelectionOrigin Is Nothing Then
                            Set moCellSelectionOrigin = moAreaSelection.moCellPosition.Copy
                        End If
                    Else
                        Set moCellSelectionOrigin = moAreaSelection.moCellPosition.Copy
                    End If
            End Select
            
            Select Case moAreaSelection.mhaAreaSelection
                Case haCell
                    lTextPosition = moAreaSelection.moCellPosition.TextPosition
                    
                    If moSelectionn.SelectionType = stText Then
                        If lTextPosition >= moSelectionn.StartPosition And lTextPosition <= moSelectionn.EndPosition Then
                            tmrDrag.Enabled = True
                        End If
                    End If
            End Select
    End Select
End Sub

Private Sub tmrDrag_Timer()
    Me.MousePointer = vbCustom
    Me.MouseIcon = moDragPointer
    mcsCursorState = csDragText
    tmrDrag.Enabled = False
End Sub

Private Sub goCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bLeftShift As Boolean
    Dim bRightShift As Boolean
    Dim bLeftControl As Boolean
    Dim bRightControl As Boolean
    Dim bLeftAlt As Boolean
    Dim bRightAlt As Boolean
    Dim lTextPosition As Long
    
    bLeftShift = GetKeyState(VK_LSHIFT) < 0
    bRightShift = GetKeyState(VK_RSHIFT) < 0
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    bRightControl = GetKeyState(VK_RCONTROL) < 0
    bLeftAlt = GetKeyState(VK_LMENU) < 0
    bRightAlt = GetKeyState(VK_RMENU) < 0
    
    If tmrDrag.Enabled = True Then
        Me.MousePointer = vbCustom
        Me.MouseIcon = moDragPointer
        mcsCursorState = csDragText
        tmrDrag.Enabled = False
    End If

    moAreaSelection.GetSelection X, Y
    Select Case moAreaSelection.mhaAreaSelection
        Case haColumnSelector
            MousePointer = vbCustom
            MouseIcon = moDownPointer
        Case haRowSelector
            MousePointer = vbCustom
            MouseIcon = moRightPointer
        Case haTableSelector
            MousePointer = vbDefault
        Case haCell
            If bLeftControl Or bRightControl Then
                MousePointer = vbCustom
                MouseIcon = moSelectCellPointer
            Else
                If mcsCursorState = csDragText Then
                    MousePointer = vbCustom
                    MouseIcon = moDragPointer
                Else
                    MousePointer = vbIbeam
                End If
            End If
        Case haOther
            MousePointer = vbDefault
    End Select
    

    Select Case Button
        Case vbLeftButton
            Select Case moAreaSelection.mhaAreaSelection
                Case haColumnSelector
                    If moAreaSelection.moCellPosition.ColumnIndex > -1 Then
                        moSelectionn.ColumnRange(moCellSelectionOrigin.ColumnIndex, moAreaSelection.moCellPosition.ColumnIndex) = True
                        Rerender False, False
                    End If
                Case haRowSelector
                    If moAreaSelection.moCellPosition.RowIndex > -1 Then
                        moSelectionn.RowRange(moCellSelectionOrigin.RowIndex, moAreaSelection.moCellPosition.RowIndex) = True
                        Rerender False, False
                    End If
                Case haCell
                    Select Case moSelectionn.SelectionType
                        Case stCells
                            moSelectionn.Range(moCellSelectionOrigin.RowIndex, moCellSelectionOrigin.ColumnIndex, moAreaSelection.moCellPosition.RowIndex, moAreaSelection.moCellPosition.ColumnIndex) = True
                            Rerender False
                        Case stText
                            Select Case mcsCursorState
                                Case csNone
                                    moSelectionn.StartPosition = moCellSelectionOrigin.TextPosition
                                    moSelectionn.EndPosition = moAreaSelection.moCellPosition.TextPosition
                                    moCellPosition.CopyOf moAreaSelection.moCellPosition
                                    Rerender False
                                Case csDragText
                                    moCellPosition.CopyOf moAreaSelection.moCellPosition
                                    Rerender False
                            End Select
                    End Select
            End Select
        Case 0
            Select Case moAreaSelection.mhaAreaSelection
                Case haCell
                    lTextPosition = moAreaSelection.moCellPosition.TextPosition
                    
                    If moSelectionn.SelectionType = stText Then
                        If lTextPosition >= moSelectionn.StartPosition And lTextPosition <= moSelectionn.EndPosition Then
                            Me.MousePointer = vbDefault
                        End If
                    End If
            End Select
    End Select
End Sub

Private Sub goCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bLeftShift As Boolean
    Dim bRightShift As Boolean
    Dim bLeftControl As Boolean
    Dim bRightControl As Boolean
    Dim bLeftAlt As Boolean
    Dim bRightAlt As Boolean
    
    bLeftShift = GetKeyState(VK_LSHIFT) < 0
    bRightShift = GetKeyState(VK_RSHIFT) < 0
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    bRightControl = GetKeyState(VK_RCONTROL) < 0
    bLeftAlt = GetKeyState(VK_LMENU) < 0
    bRightAlt = GetKeyState(VK_RMENU) < 0
    
    tmrDrag.Enabled = False
    
    mmuvMouse.Button = Button
    mmuvMouse.X = X
    mmuvMouse.Y = Y
    mmuvMouse.UpDown = 1
    
    Select Case Button
        Case vbLeftButton
            moAreaSelection.GetSelection X, Y
            Select Case moAreaSelection.mhaAreaSelection
                Case haColumnSelector
                    If bLeftControl Or bRightControl Then
                        moSelectionn.Column(moAreaSelection.moCellPosition.ColumnIndex) = Not moSelectionn.Column(moAreaSelection.moCellPosition.ColumnIndex)
                    Else
                        moSelectionn.ColumnRange(moCellSelectionOrigin.ColumnIndex, moAreaSelection.moCellPosition.ColumnIndex) = True
                    End If
                    Rerender False, False
                    moCursor.RecreateCursor
                Case haRowSelector
                    If bLeftControl Or bRightControl Then
                        moSelectionn.Row(moAreaSelection.moCellPosition.RowIndex) = Not moSelectionn.Column(moAreaSelection.moCellPosition.RowIndex)
                    Else
                        moSelectionn.RowRange(moCellSelectionOrigin.RowIndex, moAreaSelection.moCellPosition.RowIndex) = True
                    End If
                    Rerender False, False
                    moCursor.RecreateCursor
                Case haTableSelector
                    moSelectionn.SelectionType = stTable
                    Rerender False
                Case haOther
                    moSelectionn.SelectionType = stNone
                    Rerender False
                Case haCell
                    Select Case moSelectionn.SelectionType
                        Case stCells
                            moSelectionn.Range(moCellSelectionOrigin.RowIndex, moCellSelectionOrigin.ColumnIndex, moAreaSelection.moCellPosition.RowIndex, moAreaSelection.moCellPosition.ColumnIndex) = True
                            Rerender False
                        Case stText
                            Select Case mcsCursorState
                                Case csNone
                                    moSelectionn.StartPosition = moCellSelectionOrigin.TextPosition
                                    moSelectionn.EndPosition = moAreaSelection.moCellPosition.TextPosition
                                    moCellPosition.CopyOf moAreaSelection.moCellPosition
                                    If moSelectionn.StartPosition = moSelectionn.EndPosition Then
                                        moSelectionn.SelectionType = stNone
                                    End If
                                    Rerender False
                                Case csDragText
                                    Dim sCopyText As String
                                    Dim lTextPosition As Long
                                    
                                    mcsCursorState = csNone
                                    
                                    lTextPosition = moAreaSelection.moCellPosition.TextPosition
                                    
                                    If lTextPosition >= moSelectionn.StartPosition And lTextPosition <= moSelectionn.EndPosition Then
                                    Else
                                        sCopyText = Mid$(moTable.Text, moSelectionn.StartPosition + 1, moSelectionn.EndPosition - moSelectionn.StartPosition)
                                        If lTextPosition > moSelectionn.EndPosition Then
                                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                                            moTable.Text = Left$(moTable.Text, moSelectionn.StartPosition) & Mid$(moTable.Text, moSelectionn.EndPosition + 1, lTextPosition - moSelectionn.EndPosition) & sCopyText & Mid$(moTable.Text, lTextPosition + 1)
                                            moSelectionn.StartPosition = lTextPosition - Len(sCopyText)
                                            moSelectionn.EndPosition = lTextPosition
                                            moCellPosition.TextPosition = moSelectionn.EndPosition
                                            Rerender True, True
                                        Else
                                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                                            moTable.Text = Left$(moTable.Text, lTextPosition) & sCopyText & Mid$(moTable.Text, lTextPosition + 1, moSelectionn.StartPosition - lTextPosition) & Mid$(moTable.Text, moSelectionn.EndPosition + 1)
                                            moSelectionn.StartPosition = lTextPosition
                                            moSelectionn.EndPosition = lTextPosition + Len(sCopyText)
                                            moCellPosition.TextPosition = moSelectionn.EndPosition
                                            Rerender True, True
                                        End If
                                    End If
                            End Select
                    End Select
            End Select
    End Select
End Sub

Private Sub goCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lCharacterCount As Long
    
    Dim bShift As Boolean
    Dim bControl As Boolean
    Dim bLeftShift As Boolean
    Dim bRightShift As Boolean
    Dim bLeftControl As Boolean
    Dim bRightControl As Boolean
    Dim bLeftAlt As Boolean
    Dim bRightAlt As Boolean
    
    Dim yKeyState(0 To 255) As Byte
    Dim sKey  As String * 2
    Dim lScanCode As Long
    Dim msCopyText As String
    Dim vColumn As Variant
    Dim vRow As Variant
    Dim vChange As Variant
    Dim lRow As Long
    Dim lColumn As Long
    Dim lStartRow As Long
    Dim lEndRow As Long
    Dim oStartCell As CellPosition
    Dim oEndCell As CellPosition
    
    bLeftShift = GetKeyState(VK_LSHIFT) < 0
    bRightShift = GetKeyState(VK_RSHIFT) < 0
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    bRightControl = GetKeyState(VK_RCONTROL) < 0
    bLeftAlt = GetKeyState(VK_LMENU) < 0
    bRightAlt = GetKeyState(VK_RMENU) < 0

    bShift = bLeftShift Or bRightShift
    bControl = bLeftControl Or bRightControl
    
    lScanCode = MapVirtualKey(CLng(KeyCode), MAPVK_VK_TO_VSC)
    GetKeyboardState yKeyState(0)
    ToAscii CLng(KeyCode), lScanCode, yKeyState(0), sKey, 0
        
    Select Case Left$(sKey, 1)
        Case " "
            If bControl Then
                mnuClearCell_Click
            ElseIf bShift Then
                moChangeHistory.Clear
                Me.scrVertical.Value = 0
                moTable.Text = ""
                moCellPosition.TextPosition = 0
                moSelectionn.SelectionType = stNone
                Rerender
            Else
                DeleteSelection
                moTable.InsertText Chr$(KeyCode), moCellPosition
                Rerender
            End If
        Case Is > " "
            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
            DeleteSelection
            moTable.InsertText Left$(sKey, 1), moCellPosition
            Rerender
        Case Else
            Select Case KeyCode
                Case vbKeyEnd
                    moSelectionn.SelectionType = stNone
                    If bShift And bControl Then
                        If moSelectionn.SelectionType <> stText Then
                            moSelectionn.StartPosition = moCellPosition.TextPosition
                        End If
                        moSelectionn.SelectionType = stText
                        moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex))
                        moSelectionn.EndPosition = moCellPosition.TextPosition
                        Rerender
                    ElseIf bShift Then
                        If moSelectionn.SelectionType <> stText Then
                            moSelectionn.StartPosition = moCellPosition.TextPosition
                        End If
                        moSelectionn.SelectionType = stText
                        moCellPosition.ColumnIndex = UBound(moTable.TableRow(moCellPosition.RowIndex))
                        moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex))
                        moSelectionn.EndPosition = moCellPosition.TextPosition
                        Rerender
                    ElseIf bControl Then
                        moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex))
                        Rerender
                        UpdateCursor
                    Else
                        moCellPosition.ColumnIndex = UBound(moTable.TableRow(moCellPosition.RowIndex))
                        moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex))
                        Rerender
                        UpdateCursor
                    End If
                Case vbKeyHome
                    moSelectionn.SelectionType = stNone
                    If bShift And bControl Then
                        If moSelectionn.SelectionType <> stText Then
                            moSelectionn.StartPosition = moCellPosition.TextPosition
                        End If
                        moSelectionn.SelectionType = stText
                        moCellPosition.CellTextPosition = 0
                        Dim lDummy As Long
                        lDummy = moCellPosition.HorzPixelPosition
                        moSelectionn.EndPosition = moCellPosition.TextPosition
                        Rerender
                        UpdateCursor
                    ElseIf bShift Then
                        If moSelectionn.SelectionType <> stText Then
                            moSelectionn.StartPosition = moCellPosition.TextPosition
                        End If
                        moSelectionn.SelectionType = stText
                        moCellPosition.ColumnIndex = 0
                        moCellPosition.CellTextPosition = 0
                        lDummy = moCellPosition.HorzPixelPosition
                        moSelectionn.EndPosition = moCellPosition.TextPosition
                        Rerender
                        UpdateCursor
                    ElseIf bControl Then
                        moCellPosition.CellTextPosition = 0
                        Rerender
                        UpdateCursor
                    Else
                        moCellPosition.ColumnIndex = 0
                        moCellPosition.CellTextPosition = 0
                        lDummy = moCellPosition.HorzPixelPosition
                        Rerender
                        UpdateCursor
                    End If
                Case vbKeyF3
                    txtFind.SetFocus
                Case vbKeyInsert
                    If bLeftShift And bLeftControl Then
                        mnuCopyCell_Click
                    ElseIf bLeftShift Then
                        mnuCopyColumn_Click
                    ElseIf bLeftControl Then
                        mnuCopyRow_Click
                    ElseIf bLeftControl Or bRightControl Then
                        If moSelectionn.SelectionType = stText Then
                            If moSelectionn.EndPosition > moSelectionn.StartPosition Then
                                msCopyText = Mid$(moTable.Text, moSelectionn.StartPosition + 1, moSelectionn.EndPosition - moSelectionn.StartPosition)
                            ElseIf moSelectionn.EndPosition < moSelectionn.StartPosition Then
                                msCopyText = Mid$(moTable.Text, moSelectionn.EndPosition + 1, moSelectionn.StartPosition - moSelectionn.EndPosition)
                            End If
                            Clipboard.Clear
                            Clipboard.SetText msCopyText, vbCFText
                        End If
                    Else
                        If Clipboard.GetFormat(vbCFText) Then
                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                            DeleteSelection
                            moTable.InsertText Clipboard.GetText(vbCFText), moCellPosition
                            Rerender
                        End If
                    End If
                Case vbKeyX
                    If bControl Then
                        CopySelection
                        DeleteSelection
                        Rerender
                    End If
                Case vbKeyC
                    If bControl Then
                        CopySelection
                    End If
                Case vbKeyV
                    If Clipboard.GetFormat(vbCFText) Then
                        moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                        DeleteSelection
                        moTable.InsertText Clipboard.GetText(vbCFText), moCellPosition
                        Rerender
                    End If

                Case vbKeyZ
                    vChange = moChangeHistory.Undo
                    If Not IsEmpty(vChange) Then
                        moTable.Text = vChange(0)
                        moTable.ConvertTextToTable
                        moCellPosition.TextPosition = vChange(1)
                        'moSelectionn.Position = vChange(2)
                        moSelectionn.SelectionType = stNone
                        Rerender
                    End If
                Case vbKeyEscape
                    moSelectionn.SelectionType = stNone
                    moCursor.RecreateCursor
                    Rerender

                Case vbKeyDelete
                    If Not bControl Then
                        If moSelectionn.SelectionType = stNone Then
                            If moCellPosition.TextPosition < Len(moTable.Text) Then
                                If Mid$(moTable.Text, moCellPosition.TextPosition + 1, Len(moTable.msRowDelimiter)) = moTable.msRowDelimiter Then
                                    lCharacterCount = Len(moTable.msRowDelimiter)
                                Else
                                    lCharacterCount = 1
                                End If
                                moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                                moTable.DeleteText moCellPosition.TextPosition, lCharacterCount
                                Rerender
                            End If
                        Else
                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                            DeleteSelection bLeftControl Or bRightControl
                            'Rerender
                        End If
                    Else
                        ModifySelection mfClear
                    End If
                Case vbKeyBack
                    If Not moSelectionn.SelectionType = stText Then
                        If moCellPosition.TextPosition > 0 Then
                            If (moCellPosition.TextPosition - Len(moTable.msRowDelimiter) + 1) >= 1 Then
                                If Mid$(moTable.Text, moCellPosition.TextPosition - Len(moTable.msRowDelimiter) + 1, Len(moTable.msRowDelimiter)) = moTable.msRowDelimiter Then
                                    lCharacterCount = Len(moTable.msRowDelimiter)
                                ElseIf Mid$(moTable.Text, moCellPosition.TextPosition - Len(moTable.msColumnDelimiter) + 1, Len(moTable.msColumnDelimiter)) = moTable.msColumnDelimiter Then
                                    lCharacterCount = Len(moTable.msColumnDelimiter)
                                Else
                                    lCharacterCount = 1
                                End If
                            Else
                                lCharacterCount = 1
                            End If
                            moCellPosition.TextPosition = moCellPosition.TextPosition - lCharacterCount
                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                            moTable.DeleteText moCellPosition.TextPosition, lCharacterCount
                            Rerender
                        End If
                    Else
                        moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                        DeleteSelection
                        Rerender
                    End If
                Case vbKeyReturn
                    If bLeftShift Then
                        mnuInsertColumn_Click
                    ElseIf bLeftControl Then
                        mnuInsertRow_Click
                    Else
                        moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                        DeleteSelection
                        moTable.InsertText moTable.msRowDelimiter, moCellPosition
                        moSelectionn.SelectionType = stNone
                        Rerender
                    End If
                Case vbKeyTab
                    If bRightAlt Then
                        txtFind.SetFocus
                        txtReplace.TabStop = True
                    ElseIf bLeftControl Then
                        mnuInsertCellDown_Click
                    ElseIf bLeftShift Then
                        If Not moSelectionn.SelectionType = stText Then
                            mnuInsertCellRight_Click
                        Else
                            Set oStartCell = moCellPosition.Copy
                            Set oEndCell = moCellPosition.Copy
                            oStartCell.TextPosition = moSelectionn.StartPosition
                            oEndCell.TextPosition = moSelectionn.EndPosition
                            lStartRow = oStartCell.RowIndex
                            lEndRow = oEndCell.RowIndex

                            For lRow = lStartRow To lEndRow
                                vRow = moTable.TableRow(lRow)
                                If vRow(0) = "" Then
                                    For lColumn = 0 To UBound(vRow) - 1
                                        vRow(lColumn) = vRow(lColumn + 1)
                                    Next
                                    ReDim Preserve vRow(UBound(vRow) - 1)
                                    moTable.TableRow(lRow) = vRow
                                End If
                            Next
                            If moCellPosition.ColumnIndex > 0 Then
                                moCellPosition.ColumnIndex = moCellPosition.ColumnIndex - 1
                            End If
                            
                            oStartCell.ColumnIndex = 0
                            oStartCell.CellTextPosition = 0
                            If oEndCell.ColumnIndex > 0 Then
                                oEndCell.ColumnIndex = oEndCell.ColumnIndex - 1
                            End If
                            oStartCell.moTableInfo.moTable.Table = moTable.Table
                            oEndCell.moTableInfo.moTable.Table = moTable.Table
                            moSelectionn.SelectionType = stNone
                            'moSelectionn.Position = oStartCell.TextPosition
                            moSelectionn.SelectionType = stText
                            'moSelectionn.Position = oEndCell.TextPosition

                            Set oStartCell = Nothing
                            Set oEndCell = Nothing
                            Rerender
                        End If
                    ElseIf bRightControl Then
                        'mnuDeleteCellDown_Click
                    ElseIf bRightShift Then
                        'mnuDeleteCellRight_Click
                    Else
                        If moSelectionn.SelectionType <> stText Then
                            moSelectionn.SelectionType = stNone
                            moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
                            moTable.InsertText Chr$(KeyCode), moCellPosition
                            Rerender
                        Else
                            Set oStartCell = moCellPosition.Copy
                            Set oEndCell = moCellPosition.Copy
                            oStartCell.TextPosition = moSelectionn.StartPosition
                            oEndCell.TextPosition = moSelectionn.EndPosition
                            lStartRow = oStartCell.RowIndex
                            lEndRow = oEndCell.RowIndex

                            For lRow = lStartRow To lEndRow
                                vRow = moTable.TableRow(lRow)
                                ReDim Preserve vRow(UBound(vRow) + 1)
                                For lColumn = UBound(vRow) To 1 Step -1
                                    vRow(lColumn) = vRow(lColumn - 1)
                                Next
                                vRow(0) = ""
                                moTable.TableRow(lRow) = vRow
                            Next
                            moCellPosition.ColumnIndex = moCellPosition.ColumnIndex + 1
                            
                            oStartCell.ColumnIndex = 0
                            oStartCell.CellTextPosition = 0
                            oEndCell.ColumnIndex = oEndCell.ColumnIndex + 1
                            oStartCell.moTableInfo.moTable.Table = moTable.Table
                            oEndCell.moTableInfo.moTable.Table = moTable.Table
                            moSelectionn.SelectionType = stNone
                            'moSelectionn.Position = oStartCell.TextPosition
                            moSelectionn.SelectionType = stText
                            'moSelectionn.Position = oEndCell.TextPosition

                            Set oStartCell = Nothing
                            Set oEndCell = Nothing
                            Rerender
                        End If
                    End If
                Case vbKeyRight
                    If bControl Then
                        If moSelectionn.SelectionType = stNone Or moSelectionn.SelectionType = stText Then
                            If moCellPosition.ColumnIndex < UBound(moTable.TableRow(moCellPosition.RowIndex)) Then
                                moCellPosition.ColumnIndex = moCellPosition.ColumnIndex + 1
                            End If
                            moCellPosition.CellTextPosition = 0
                            UpdateCursor
                            Rerender
                        Else
                            MoveSelection mdRight
                        End If
                    Else
                        If bShift Then
                            If moSelectionn.SelectionType <> stText Then
                                moSelectionn.StartPosition = moCellPosition.TextPosition
                            End If
                            moSelectionn.SelectionType = stText
                        Else
                            moSelectionn.SelectionType = stNone
                        End If
                        
                        moCellPosition.TextPosition = moCellPosition.TextPosition + 1
                        If moCellPosition.TextPosition <= Len(moTable.Text) Then
                            If Mid$(moTable.Text, moCellPosition.TextPosition, 2) = moTable.msRowDelimiter Then
                                moCellPosition.TextPosition = moCellPosition.TextPosition + Len(moTable.msRowDelimiter) - 1
                            End If
                        End If
                        If moCellPosition.TextPosition > Len(moTable.Text) Then
                            moCellPosition.TextPosition = 0
                        End If
                        
                        If bShift Then
                            moSelectionn.EndPosition = moCellPosition.TextPosition
                        End If
                        
                        Rerender False
                    End If
                Case vbKeyLeft
                    If bControl Then
                        If moSelectionn.SelectionType = stNone Or moSelectionn.SelectionType = stText Then
                            If moCellPosition.ColumnIndex > 0 Then
                                moCellPosition.ColumnIndex = moCellPosition.ColumnIndex - 1
                            End If
                            moCellPosition.CellTextPosition = 0
                            UpdateCursor
                            Rerender
                        Else
                            MoveSelection mdLeft
                        End If
                    Else
                        If bShift Then
                            If moSelectionn.SelectionType <> stText Then
                                moSelectionn.StartPosition = moCellPosition.TextPosition
                            End If
                            moSelectionn.SelectionType = stText
                        Else
                            moSelectionn.SelectionType = stNone
                        End If
                        
                        moCellPosition.TextPosition = moCellPosition.TextPosition - 1
                        If moCellPosition.TextPosition >= 0 Then
                            If Mid$(moTable.Text, moCellPosition.TextPosition + 1, 1) = vbLf Then
                                moCellPosition.TextPosition = moCellPosition.TextPosition - 1
                            End If
                        End If
                        If moCellPosition.TextPosition < 0 Then
                            moCellPosition.TextPosition = Len(moTable.Text)
                        End If
                        
                        If bShift Then
                            moSelectionn.EndPosition = moCellPosition.TextPosition
                        End If
                        
                        Rerender False
                    End If
                Case vbKeyUp
                    If bControl Then
                        MoveSelection mdUp
                    Else
                        If bShift Then
                            If moSelectionn.SelectionType <> stText Then
                                moSelectionn.StartPosition = moCellPosition.TextPosition
                            End If
                            moSelectionn.SelectionType = stText
                        Else
                            moSelectionn.SelectionType = stNone
                        End If
                        
                        moCellPosition.RowIndex = moCellPosition.RowIndex - 1
                        If moCellPosition.RowIndex < 0 Then
                            moCellPosition.RowIndex = UBound(moTable.Table)
                        End If
                        If moCellPosition.ColumnIndex > UBound(moTable.TableRow(moCellPosition.RowIndex)) Then
                            moCellPosition.ColumnIndex = UBound(moTable.TableRow(moCellPosition.RowIndex))
                            If moCellPosition.ColumnIndex > -1 Then
                                moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex)) + 1
                            Else
                                moCellPosition.ColumnIndex = 0
                                moCellPosition.CellTextPosition = 0
                            End If
                        End If
                        
                        If bShift Then
                            moSelectionn.EndPosition = moCellPosition.TextPosition
                        End If
                        
                        Rerender False
                    End If
                Case vbKeyDown
                    If bControl Then
                        MoveSelection mdDown
                    Else
                        If bShift Then
                            If moSelectionn.SelectionType <> stText Then
                                moSelectionn.StartPosition = moCellPosition.TextPosition
                            End If
                            moSelectionn.SelectionType = stText
                        Else
                            moSelectionn.SelectionType = stNone
                        End If
                        
                        moCellPosition.RowIndex = moCellPosition.RowIndex + 1
                        If moCellPosition.RowIndex > UBound(moTable.Table) Then
                            moCellPosition.RowIndex = 0
                        End If
                        If moCellPosition.ColumnIndex > UBound(moTable.TableRow(moCellPosition.RowIndex)) Then
                            moCellPosition.ColumnIndex = UBound(moTable.TableRow(moCellPosition.RowIndex))
                            If moCellPosition.ColumnIndex > -1 Then
                                moCellPosition.CellTextPosition = Len(moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex)) + 1
                            Else
                                moCellPosition.ColumnIndex = 0
                                moCellPosition.CellTextPosition = 0
                            End If
                        End If
                        Dim sCell As String
                        
                        sCell = moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex)
                        If moCellPosition.CellTextPosition > Len(sCell) Then
                            moCellPosition.CellTextPosition = Len(sCell)
                        End If
                        If bShift Then
                            moSelectionn.EndPosition = moCellPosition.TextPosition
                        End If
                        Rerender False
                    End If
            End Select
    End Select
End Sub



Private Sub mnuClearCell_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.ClearCell moCellPosition.RowIndex, moCellPosition.ColumnIndex
    moCellPosition.CellTextPosition = 0
    Rerender
End Sub

Private Sub mnuCopyCell_Click()
    Clipboard.Clear
    Clipboard.SetText moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex)
End Sub

Private Sub mnuCopyColumn_Click()
    Clipboard.Clear
    Clipboard.SetText Join(moTable.ArrayFromTableColumn(moCellPosition.ColumnIndex), moTable.msRowDelimiter)
End Sub

Private Sub mnuCopyRow_Click()
    Clipboard.Clear
    Clipboard.SetText Join(moTable.ArrayFromTableRow(moCellPosition.ColumnIndex), moTable.msColumnDelimiter)
End Sub


Private Sub mnuHTMLFloated_Click()
    Dim vRows As Variant
    Dim vRow As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    vRows = Array()
    For Each vRow In moTable.Table
        ArrayAppend vRows, "<div id="""">" & Join(vRow, "</div><div id="""">") & "</div>"
    Next
    moTable.Text = Join(vRows, vbCrLf)
    moSelectionn.SelectionType = stNone
    moCellPosition.TextPosition = 0
    Rerender
End Sub

Private Sub mnuHTMLTable_Click()
    Dim vRows As Variant
    Dim vRow As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    vRows = Array()
    For Each vRow In moTable.Table
        ArrayAppend vRows, "<td>" & Join(vRow, "</td><td>") & "</td>"
    Next
    moTable.Text = "<table>" & vbCrLf & vbTab & "<tr>" & vbCrLf & vbTab & vbTab & Join(vRows, vbCrLf & vbTab & "</tr>" & vbCrLf & vbTab & "<tr>" & vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</tr>" & vbCrLf & "</table>"
    moSelectionn.SelectionType = stNone
    moCellPosition.TextPosition = 0
    Rerender
End Sub

Private Sub mnuInsertColumn_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.InsertBlankColumn moCellPosition.ColumnIndex
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    Rerender
End Sub

Private Sub mnuInsertColumnRight_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.InsertBlankColumn moCellPosition.ColumnIndex + 1
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    moCellPosition.ColumnIndex = moCellPosition.ColumnIndex + 1
    Rerender
End Sub

Private Sub mnuInsertPasteRight_Click()
    Dim vClipboard1 As Variant
    Dim vClipboard2 As Variant
    Dim lRow As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lTableHeight As Long
    Dim vTable As Variant
    Dim vPad As Variant
    
    If Clipboard.GetFormat(vbCFText) Then
        moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
        DeleteSelection

        vClipboard1 = Split(Clipboard.GetText(vbCFText), moTable.msRowDelimiter)
        lHeight = UBound(vClipboard1)
        ReDim vClipboard2(lHeight)
        For lRow = 0 To lHeight
            vClipboard2(lRow) = Split(vClipboard1(lRow), moTable.msColumnDelimiter)
            If UBound(vClipboard2(lRow)) > lWidth Then
                lWidth = UBound(vClipboard2(lRow))
            End If
        Next
        
        vTable = moTable.Table
        lTableHeight = UBound(vTable)
        If (moCellPosition.RowIndex + lHeight) > lTableHeight Then
            ReDim Preserve vTable(moCellPosition.RowIndex + lHeight)
            For lRow = lTableHeight + 1 To moCellPosition.RowIndex + lHeight
                vTable(lRow) = Array("")
            Next
        End If
        
        For lRow = moCellPosition.RowIndex To moCellPosition.RowIndex + lHeight
            If UBound(vTable(lRow)) < moCellPosition.ColumnIndex Then
                vTable(lRow) = Concat(vTable(lRow), PaddedArray(moCellPosition.ColumnIndex - UBound(vTable(lRow)) - 2, ""), vClipboard2(lRow - moCellPosition.RowIndex))
            Else
                vTable(lRow) = Concat(Splice(vTable(lRow), 0, moCellPosition.ColumnIndex - 1), vClipboard2(lRow - moCellPosition.RowIndex), PaddedArray(lWidth - UBound(vClipboard2(lRow - moCellPosition.RowIndex)) - 1, ""), Splice(vTable(lRow), moCellPosition.ColumnIndex))
            End If
        Next
        moTable.Table = vTable
        moSelectionn.SelectionType = stNone
        Rerender
    End If
End Sub

Private Sub mnuInsertPasteDown_Click()
    Dim vClipboard1 As Variant
    Dim vClipboard2 As Variant
    Dim lRow As Long
    Dim lHeight As Long
    Dim lWidth As Long
    
    If Clipboard.GetFormat(vbCFText) Then
        moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
        DeleteSelection

        vClipboard1 = Split(Clipboard.GetText(vbCFText), moTable.msRowDelimiter)
        lHeight = UBound(vClipboard1)
        ReDim vClipboard2(lHeight)
        For lRow = 0 To lHeight
            vClipboard2(lRow) = Split(vClipboard1(lRow), moTable.msColumnDelimiter)
            If UBound(vClipboard2(lRow)) > lWidth Then
                lWidth = UBound(vClipboard2(lRow))
            End If
        Next
        
'        For lRow = moCellPosition.RowIndex To moCellPosition.RowIndex + lHeight
'            If UBound(vTable(lRow)) < moCellPosition.ColumnIndex Then
'                vTable(lRow) = Concat(vTable(lRow), PaddedArray(moCellPosition.ColumnIndex - UBound(vTable(lRow)) - 2, ""), vClipboard2(lRow - moCellPosition.RowIndex))
'            Else
'                vTable(lRow) = Concat(Splice(vTable(lRow), 0, moCellPosition.ColumnIndex - 1), vClipboard2(lRow - moCellPosition.RowIndex), PaddedArray(lWidth - UBound(vClipboard2(lRow - moCellPosition.RowIndex)) - 1, ""), Splice(vTable(lRow), moCellPosition.ColumnIndex))
'            End If
'        Next
'        moTable.Table = vTable
'        moSelectionn.SelectionType=stNone
'        Rerender
    End If
End Sub

Private Sub mnuInsertRow_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.Table = Concat(Splice(moTable.Table, 0, moCellPosition.RowIndex - 1), Array(Array("")), Splice(moTable.Table, moCellPosition.RowIndex))
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    Rerender
End Sub

Private Sub mnuInsertCellDown_Click()
    Dim vColumn As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    vColumn = moTable.ArrayFromTableColumn(moCellPosition.ColumnIndex)
    vColumn = Concat(Splice(vColumn, 0, moCellPosition.RowIndex - 1), Array(""), Splice(vColumn, moCellPosition.RowIndex))
    moTable.ArrayIntoTableColumn moCellPosition.ColumnIndex, vColumn
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    Rerender
End Sub

Private Sub mnuInsertCellRight_Click()
    Dim vRow As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    vRow = moTable.ArrayFromTableRow(moCellPosition.RowIndex)
    vRow = Concat(Splice(vRow, 0, moCellPosition.ColumnIndex - 1), Array(""), Splice(vRow, moCellPosition.ColumnIndex))
    moTable.ArrayIntoTableRow moCellPosition.RowIndex, vRow
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    Rerender
End Sub

Private Sub mnuInsertRowDown_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.Table = Concat(Splice(moTable.Table, 0, moCellPosition.RowIndex), Array(Array("")), Splice(moTable.Table, moCellPosition.RowIndex + 1))
    moSelectionn.SelectionType = stNone
    moCellPosition.CellTextPosition = 0
    moCellPosition.RowIndex = moCellPosition.RowIndex + 1
    Rerender
End Sub

Private Sub mnuSelectCell_Click()
    moSelectionn.SelectionType = stCells
    moSelectionn.Cell(moCellPosition.RowIndex, moCellPosition.ColumnIndex) = True
    Rerender False
End Sub

Private Sub mnuSelectColumn_Click()
    moSelectionn.SelectionType = stColumns
    moSelectionn.Column(moCellPosition.ColumnIndex) = True
    Rerender
End Sub

Private Sub mnuSelectionCamelCase_Click()
    ModifySelection mfCamelCase
End Sub

Private Sub mnuSelectionClear_Click()
    ModifySelection mfClear
End Sub

Private Sub mnuSelectionCommaList_Click()
    CommaList
End Sub

Private Sub mnuSelectionCommaListQuotes_Click()
    CommaList True
End Sub

Private Sub CommaList(Optional ByVal bWithQuotes As Boolean = False)
    Dim lRow As Long
    Dim lColumn As Long
    Dim sText As String
    Dim vTable As Variant
    Dim sDelimiter As String
    Dim sQuote As String
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    If bWithQuotes Then
        sDelimiter = "','"
        sQuote = "'"
    Else
        sDelimiter = ","
        sQuote = ""
    End If
    
    Select Case moSelectionn.SelectionType
        Case stNone
            For lRow = 0 To moTable.LastRow
                For lColumn = 0 To UBound(moTable.TableRow(lRow))
                    sText = sText & sDelimiter & moTable.TableCell(lRow, lColumn)
                Next
            Next
            moTable.Text = sQuote & Mid$(sText, 2) & sQuote
            If scrVertical.Value > moTable.LastRow Then
                scrVertical.Value = moTable.LastRow
            End If
            Rerender
        Case stColumns
            ReDim vTable(moTable.LastRow)
            For lRow = 0 To moTable.LastRow
                For lColumn = 0 To UBound(moTable.TableRow(lRow))
                    
                Next
            Next
        Case stRows
            ReDim vTable(moTable.LastRow)
            For lRow = 0 To moTable.LastRow
                If moSelectionn.Row(lRow) Then
                    vTable(lRow) = Array(sQuote & Join(moTable.TableRow(lRow), sDelimiter) & sQuote)
                Else
                    vTable(lRow) = moTable.TableRow(lRow)
                End If
            Next
            moTable.Table = vTable
            If scrVertical.Value > moTable.LastRow Then
                scrVertical.Value = moTable.LastRow
            End If
            Rerender
    End Select
End Sub

Private Sub mnuSelectionCopy_Click()
    CopySelection
End Sub

Private Sub mnuSelectionDelete_Click()
    DeleteSelection
End Sub


Private Sub MoveSelection(ByVal mdDirection As MoveDirections)
    Dim lRow As Long
    Dim lColumn As Long
    Dim vRow As Variant
    Dim vTable As Variant
    Dim vTableRow As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    Select Case moSelectionn.SelectionType
        Case stColumns
            Select Case mdDirection
                Case mdLeft
                    If Not moSelectionn.Column(0) Then
                        ReDim vTable(moTable.LastRow)
                        For lRow = 0 To moTable.LastRow
                            vRow = Array()
                            vTableRow = moTable.TableRow(lRow)
                            ReDim vRow(UBound(vTableRow))
                            For lColumn = 0 To UBound(vTableRow)
                                If lColumn < UBound(vTableRow) Then
                                    If moSelectionn.Column(lColumn + 1) Then
                                        vRow(lColumn) = vTableRow(lColumn + 1)
                                        vTableRow(lColumn + 1) = vTableRow(lColumn)
                                    Else
                                        vRow(lColumn) = vTableRow(lColumn)
                                    End If
                                Else
                                    vRow(lColumn) = vTableRow(lColumn)
                                End If
                            Next
                            vTable(lRow) = vRow
                        Next
                        For lColumn = 0 To moTable.LastColumn - 1
                            moSelectionn.Column(lColumn) = moSelectionn.Column(lColumn + 1)
                        Next
                        moSelectionn.Column(moTable.LastColumn) = False
                        moTable.Table = vTable
                        Rerender
                    End If
                Case mdRight
                    If Not moSelectionn.Column(moTable.LastColumn) Then
                        ReDim vTable(moTable.LastRow)
                        For lRow = 0 To moTable.LastRow
                            vRow = Array()
                            vTableRow = moTable.TableRow(lRow)
                            ReDim vRow(UBound(vTableRow))
                            For lColumn = UBound(vTableRow) To 0 Step -1
                                If lColumn > 0 Then
                                    If moSelectionn.Column(lColumn - 1) Then
                                        vRow(lColumn) = vTableRow(lColumn - 1)
                                        vTableRow(lColumn - 1) = vTableRow(lColumn)
                                    Else
                                        vRow(lColumn) = vTableRow(lColumn)
                                    End If
                                Else
                                    vRow(lColumn) = vTableRow(lColumn)
                                End If
                            Next
                            vTable(lRow) = vRow
                        Next
                        For lColumn = moTable.LastColumn To 1 Step -1
                            moSelectionn.Column(lColumn) = moSelectionn.Column(lColumn - 1)
                        Next
                        moSelectionn.Column(0) = False
                        moTable.Table = vTable
                        Rerender
                    End If
            End Select
        Case stRows
            Select Case mdDirection
                Case mdUp
                    If Not moSelectionn.Row(0) Then
                        vTable = Array()
                        ReDim vTable(moTable.LastRow)
                        For lRow = 0 To moTable.LastRow
                            If lRow < moTable.LastRow Then
                                If Not moSelectionn.Row(lRow + 1) Then
                                    vTable(lRow) = moTable.TableRow(lRow)
                                Else
                                    vTable(lRow) = moTable.TableRow(lRow + 1)
                                    moTable.TableRow(lRow + 1) = moTable.TableRow(lRow)
                                End If
                            Else
                                vTable(lRow) = moTable.TableRow(lRow)
                            End If
                        Next
                        For lRow = 0 To moTable.LastRow - 1
                            moSelectionn.Row(lRow) = moSelectionn.Row(lRow + 1)
                        Next
                        moSelectionn.Row(moTable.LastRow) = False
                        moTable.Table = vTable
                        Rerender
                    End If
                Case mdDown
                    If Not moSelectionn.Row(moTable.LastRow) Then
                        vTable = Array()
                        ReDim vTable(moTable.LastRow)
                        For lRow = moTable.LastRow To 0 Step -1
                            If lRow > 0 Then
                                If Not moSelectionn.Row(lRow - 1) Then
                                    vTable(lRow) = moTable.TableRow(lRow)
                                Else
                                    vTable(lRow) = moTable.TableRow(lRow - 1)
                                    moTable.TableRow(lRow - 1) = moTable.TableRow(lRow)
                                End If
                            Else
                                vTable(lRow) = moTable.TableRow(lRow)
                            End If
                        Next
                        For lRow = moTable.LastRow To 1 Step -1
                            moSelectionn.Row(lRow) = moSelectionn.Row(lRow - 1)
                        Next
                        moSelectionn.Row(0) = False
                        moTable.Table = vTable
                        Rerender
                    End If
            End Select
        Case stCells
            Select Case mdDirection
                Case mdUp
                Case mdDown
                Case mdLeft
                Case mdRight
            End Select
    End Select
    
    Rerender
End Sub

Private Sub DeleteSelection(Optional ByVal bDown As Boolean)
    Dim vRow As Variant
    Dim vTableRow As Variant
    Dim vTable As Variant
    Dim lRow As Long
    Dim lColumn As Long
    Dim lShift As Long
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    Select Case moSelectionn.SelectionType
        Case stColumns
            If Not moSelectionn.Column(moCellPosition.ColumnIndex) Then
                lShift = 0
                For lColumn = 0 To moCellPosition.ColumnIndex - 1
                    If moSelectionn.Column(lColumn) Then
                        lShift = lShift + 1
                    End If
                Next
                moCellPosition.ColumnIndex = moCellPosition.ColumnIndex - lShift
            Else
                moCellPosition.CellTextPosition = 0
            End If
            
            vTable = Array()
            ReDim vTable(moTable.LastRow)
            
            Dim bAllRowsBlank As Boolean
            bAllRowsBlank = True
            
            For lRow = 0 To moTable.LastRow
                vRow = Array()
                vTableRow = moTable.TableRow(lRow)
                For lColumn = 0 To UBound(vTableRow)
                    If Not moSelectionn.Column(lColumn) Then
                        ReDim Preserve vRow(UBound(vRow) + 1)
                        vRow(UBound(vRow)) = vTableRow(lColumn)
                    End If
                Next
                vTable(lRow) = Join(vRow, moTable.msColumnDelimiter)
                If UBound(vRow) >= 0 Then
                    bAllRowsBlank = False
                End If
            Next
            If Not bAllRowsBlank Then
                moTable.Text = Join(vTable, moTable.msRowDelimiter)
            Else
                moTable.Text = ""
            End If
            moSelectionn.SelectionType = stNone
            Rerender True, False
        Case stRows
            If Not moSelectionn.Row(moCellPosition.RowIndex) Then
                lShift = 0
                For lRow = 0 To moCellPosition.RowIndex - 1
                    If moSelectionn.Row(lRow) Then
                        lShift = lShift + 1
                    End If
                Next
                moCellPosition.RowIndex = moCellPosition.RowIndex - lShift
            Else
                moCellPosition.CellTextPosition = 0
            End If
        
            vTable = Array()
            For lRow = 0 To moTable.LastRow
                If Not moSelectionn.Row(lRow) Then
                    ReDim Preserve vTable(UBound(vTable) + 1)
                    vTable(UBound(vTable)) = moTable.TableRow(lRow)
                End If
            Next
            moTable.Table = vTable
            moSelectionn.SelectionType = stNone
            Rerender True, False
        Case stCells
            If Not bDown Then
                vTable = Array()
                For lRow = 0 To moTable.LastRow
                    vRow = Array()
                    vTableRow = moTable.TableRow(lRow)
                    For lColumn = 0 To UBound(vTableRow)
                        If Not moSelectionn.Cell(lRow, lColumn) Then
                            ReDim Preserve vRow(UBound(vRow) + 1)
                            vRow(UBound(vRow)) = vTableRow(lColumn)
                        End If
                    Next
                    ReDim Preserve vTable(UBound(vTable) + 1)
                    vTable(UBound(vTable)) = vRow
                Next
                moTable.Table = vTable
                moSelectionn.SelectionType = stNone
            End If
            Rerender
        Case stTable
            moTable.Text = ""
            moSelectionn.SelectionType = stNone
            moCursor.moPosition.TextPosition = 0
            Rerender
        Case stText
            If moSelectionn.EndPosition > moSelectionn.StartPosition Then
                moTable.DeleteText moSelectionn.StartPosition, moSelectionn.EndPosition - moSelectionn.StartPosition
                moCellPosition.TextPosition = moSelectionn.StartPosition
            End If
            moSelectionn.SelectionType = stNone
            Rerender
    End Select
End Sub


Private Sub mnuSelectionFill_Click()
    Dim lColumn As Long
    Dim lRow As Long
    Dim sCellText As String
    Dim bOK As Boolean
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    sCellText = moTable.TableCell(moCellPosition.RowIndex, moCellPosition.ColumnIndex)

    For lRow = 0 To moTable.LastRow
        For lColumn = 0 To moTable.LastColumn
            bOK = False
            Select Case moSelectionn.SelectionType
                Case stNone
                    bOK = True
                Case stColumns
                    bOK = moSelectionn.Column(lColumn)
                Case stRows
                    bOK = moSelectionn.Row(lRow)
                Case stCells
                    bOK = moSelectionn.Cell(lRow, lColumn)
            End Select
            If bOK Then
                If lColumn <= UBound(moTable.TableRow(lRow)) Then
                    moTable.TableCell(lRow, lColumn) = sCellText
                End If
            End If
        Next
    Next
    
    Rerender
End Sub

Private Sub mnuSelectionLowerCase_Click()
    ModifySelection mfLowerCase
End Sub

Private Sub mnuSelectionPad_Click()
    Dim vRow As Variant
    Dim lUbound As Long
    Dim lColumn As Long
    Dim lIndex As Long
    Dim lRow As Long
    Dim lSelectedColumn As Long
    Dim lSelectedRow As Long
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    Select Case moSelectionn.SelectionType
        Case stColumns
            For lColumn = moTable.LastColumn To 0 Step -1
                If moSelectionn.Column(lColumn) Then
                    Exit For
                End If
            Next
            
            For lRow = 0 To moTable.LastRow
                vRow = moTable.TableRow(lRow)
                lUbound = UBound(vRow)
                If lUbound < lColumn Then
                    ReDim Preserve vRow(lColumn)
                    moTable.TableRow(lRow) = vRow
                End If
            Next

            Rerender
        Case stRows
            For lRow = moTable.LastRow To 0 Step -1
                If moSelectionn.Row(lRow) Then
                    Exit For
                End If
            Next
            
            vRow = moTable.TableRow(lRow)
            lUbound = UBound(vRow)
            ReDim Preserve vRow(moTable.LastColumn)
            For lIndex = lUbound + 1 To moTable.LastColumn
                vRow(lIndex) = ""
            Next
            moTable.TableRow(moCellPosition.RowIndex) = vRow
            Rerender
        Case stNone
            For lRow = 0 To moTable.LastRow
                vRow = moTable.TableRow(lRow)
                lUbound = UBound(vRow)
        
                ReDim Preserve vRow(moTable.LastColumn)
                For lColumn = lUbound + 1 To UBound(vRow)
                    vRow(lColumn) = ""
                Next
                moTable.TableRow(lRow) = vRow
            Next
            Rerender
    End Select
End Sub


Private Sub mnuSelectionSpaced_Click()
    ModifySelection mfSpaced
End Sub

Private Sub mnuSelectionUnderscored_Click()
    ModifySelection mfUnderscored
End Sub

Private Sub mnuSelectionUp_Click()
    MoveSelection mdUp
End Sub

Private Sub mnuSelectionDown_Click()
    MoveSelection mdDown
End Sub

Private Sub mnuSelectionLeft_Click()
    MoveSelection mdLeft
End Sub

Private Sub mnuSelectionRight_Click()
    MoveSelection mdRight
End Sub


Private Sub mnuSelectionUpperCase_Click()
    ModifySelection mfUpperCase
End Sub

Private Sub ModifySelection(ByVal mfModifyFunction As ModifyFunctions)
    Dim lColumn As Long
    Dim lRow As Long
    Dim bChanged As Boolean
    Dim bOK As Boolean
    Dim sText As String
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    Select Case moSelectionn.SelectionType
        Case stCells, stColumns, stRows, stNone
            For lRow = 0 To moTable.LastRow
                For lColumn = 0 To UBound(moTable.TableRow(lRow))
                    bOK = False
                    Select Case moSelectionn.SelectionType
                        Case stNone
                            bOK = True
                        Case stCells
                            bOK = moSelectionn.Cell(lRow, lColumn)
                        Case stColumns
                            bOK = moSelectionn.Column(lColumn)
                        Case stRows
                            bOK = moSelectionn.Row(lRow)
                    End Select
                    
                    If bOK Then
                        Select Case mfModifyFunction
                            Case mfUpperCase
                                moTable.TableCell(lRow, lColumn) = UCase$(moTable.TableCell(lRow, lColumn))
                                bChanged = True
                            Case mfLowerCase
                                moTable.TableCell(lRow, lColumn) = LCase$(moTable.TableCell(lRow, lColumn))
                                bChanged = True
                            Case mfCamelCase
                                moTable.TableCell(lRow, lColumn) = ApplyCamelCase(moTable.TableCell(lRow, lColumn))
                                bChanged = True
                            Case mfUnderscored
                                moTable.TableCell(lRow, lColumn) = ApplyUnderscoredCase(moTable.TableCell(lRow, lColumn))
                                bChanged = True
                            Case mfSpaced
                                moTable.TableCell(lRow, lColumn) = ApplySpaceCase(moTable.TableCell(lRow, lColumn))
                                bChanged = True
                            Case mfClear
                                moTable.TableCell(lRow, lColumn) = ""
                                bChanged = True
                        End Select
                    End If
                Next
            Next
            Rerender
        Case stText
            sText = Mid$(moTable.Text, moSelectionn.StartPosition + 1, moSelectionn.EndPosition - moSelectionn.StartPosition)
            Select Case mfModifyFunction
                Case mfUpperCase
                    sText = UCase$(sText)
                    bChanged = True
                Case mfLowerCase
                    sText = LCase$(sText)
                    bChanged = True
                Case mfCamelCase
                    sText = ApplyCamelCase(sText)
                    bChanged = True
                Case mfUnderscored
                    sText = ApplyUnderscoredCase(sText)
                    bChanged = True
                Case mfSpaced
                    sText = ApplySpaceCase(sText)
                    bChanged = True
                Case mfClear
                    sText = ""
                    bChanged = True
            End Select
            moTable.Text = Left$(moTable.Text, moSelectionn.StartPosition) & sText & Mid$(moTable.Text, moSelectionn.EndPosition + 1)
            moSelectionn.EndPosition = moSelectionn.StartPosition + Len(sText)
            moCellPosition.TextPosition = moSelectionn.EndPosition
            Rerender
    End Select
    If Not bChanged Then
        moChangeHistory.Undo
    End If
End Sub

Private Sub CopySelection()
    Dim vCopiedCells As Variant
    Dim lColumn As Long
    Dim lRow As Long
    Dim vRow As Variant
    Dim vCopiedRows As Variant
    Dim sCopyText As String
    Dim bOK As Boolean
    
    vCopiedCells = Array()
    
    Select Case moSelectionn.SelectionType
        Case stCells
            vCopiedRows = Array()
            For lRow = 0 To moTable.LastRow
                bOK = False
                vCopiedCells = Array()
                For lColumn = 0 To UBound(moTable.TableRow(lRow))
                    If moSelectionn.Cell(lRow, lColumn) Then
                        bOK = True
                        
                        ReDim Preserve vCopiedCells(UBound(vCopiedCells) + 1)
                        vCopiedCells(UBound(vCopiedCells)) = moTable.TableCell(lRow, lColumn)
                    End If
                Next
                If bOK Then
                    ReDim Preserve vCopiedRows(UBound(vCopiedRows) + 1)
                    vCopiedRows(UBound(vCopiedRows)) = Join(vCopiedCells, moTable.msColumnDelimiter)
                End If
            Next
            
            sCopyText = Join(vCopiedRows, moTable.msRowDelimiter)
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
        Case stText
            sCopyText = Mid$(moTable.Text, moSelectionn.StartPosition + 1, moSelectionn.EndPosition - moSelectionn.StartPosition)
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
        Case stColumns
            vCopiedRows = Array()
            For lRow = 0 To moTable.LastRow
                vCopiedCells = Array()
                For lColumn = 0 To moTable.LastColumn
                    If moSelectionn.Column(lColumn) Then
                        If lColumn <= UBound(moTable.TableRow(lRow)) Then
                            ReDim Preserve vCopiedCells(UBound(vCopiedCells) + 1)
                            vCopiedCells(UBound(vCopiedCells)) = moTable.TableCell(lRow, lColumn)
                        End If
                    End If
                Next
                ReDim Preserve vCopiedRows(UBound(vCopiedRows) + 1)
                vCopiedRows(UBound(vCopiedRows)) = Join(vCopiedCells, moTable.msColumnDelimiter)
            Next
            sCopyText = Join(vCopiedRows, moTable.msRowDelimiter)
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
        Case stRows
            vCopiedRows = Array()
            For lRow = 0 To moTable.LastRow
                If moSelectionn.Row(lRow) Then
                    vCopiedCells = Array()
                    For lColumn = 0 To moTable.LastColumn
                        If lColumn <= UBound(moTable.TableRow(lRow)) Then
                            ReDim Preserve vCopiedCells(UBound(vCopiedCells) + 1)
                            vCopiedCells(UBound(vCopiedCells)) = moTable.TableCell(lRow, lColumn)
                        End If
                    Next
                    ReDim Preserve vCopiedRows(UBound(vCopiedRows) + 1)
                    vCopiedRows(UBound(vCopiedRows)) = Join(vCopiedCells, moTable.msColumnDelimiter)
                End If
            Next
            sCopyText = Join(vCopiedRows, moTable.msRowDelimiter)
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
        Case stTable
            sCopyText = moTable.Text
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
        Case stNone
            sCopyText = moTable.Text
            Clipboard.Clear
            Clipboard.SetText sCopyText, vbCFText
    End Select
End Sub

Private Function ApplyCamelCase(ByVal sString As String) As String
    Dim bUpper As Boolean
    Dim sChar As String
    Dim lIndex As Long
    
    bUpper = True
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        
        If sChar = "_" Then
            bUpper = True
        Else
            If bUpper Then
                ApplyCamelCase = ApplyCamelCase & UCase$(sChar)
                If LCase$(sChar) <> UCase$(sChar) Then
                    bUpper = False
                End If
            Else
                ApplyCamelCase = ApplyCamelCase & sChar
            End If
        End If
    Next
End Function

Private Function ApplyUnderscoredCase(ByVal sString As String) As String
    Dim bUnderscore As Boolean
    Dim sChar As String
    Dim sChar2 As String
    Dim lIndex As Long
    
    bUnderscore = False
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        sChar2 = Mid$(sString, lIndex + 1, 1)
        
        If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
            bUnderscore = True
        Else
            bUnderscore = False
        End If
        
        ApplyUnderscoredCase = ApplyUnderscoredCase & IIf(bUnderscore, "_", "") & LCase$(sChar)
    Next
End Function

Private Function ApplySpaceCase(ByVal sString As String) As String
    Dim bSpace As Boolean
    Dim sChar As String
    Dim sChar2 As String
    Dim lIndex As Long
    
    bSpace = False
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        sChar2 = Mid$(sString, lIndex + 1, 1)
        
        If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
            bSpace = True
        Else
            bSpace = False
        End If
        
        ApplySpaceCase = ApplySpaceCase & IIf(bSpace, " ", "") & sChar
    Next
End Function

Private Sub mnuSelectNone_Click()
    moSelectionn.SelectionType = stNone
    Rerender False
End Sub

Private Sub mnuSelectRow_Click()
    moSelectionn.SelectionType = stRows
    moSelectionn.Row(moCellPosition.RowIndex) = True
    Rerender False
End Sub



Private Sub mnuSortNumeric_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    moTable.SortByColumn moCellPosition.ColumnIndex, False, True
    Rerender
End Sub



Private Sub mnuTableHTMLList_Click()
    Dim vRows As Variant
    Dim vRow As Variant
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    vRows = Array()
    For Each vRow In moTable.Table
        ArrayAppend vRows, vbTab & "<li>" & Join(vRow, "") & "</li>"
    Next
    moTable.Text = "<ul>" & vbCrLf & Join(vRows, vbCrLf) & vbCrLf & "</ul>"
    moSelectionn.SelectionType = stNone
    moCellPosition.TextPosition = 0
    Rerender
End Sub

Private Sub mnuTableMergeColumns_Click()
    Dim lRow As Long
    Dim lDelimiterCount As Long
    
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    For lRow = 0 To UBound(moTable.Table)
        If lRow < moCellPosition.RowIndex Then
            lDelimiterCount = lDelimiterCount + UBound(moTable.TableRow(lRow)) * Len(moTable.msColumnDelimiter)
        End If
    Next
    lDelimiterCount = lDelimiterCount + moCellPosition.ColumnIndex * Len(moTable.msColumnDelimiter)
    
    moCellPosition.TextPosition = moCellPosition.TextPosition - lDelimiterCount
    
    For lRow = 0 To UBound(moTable.Table)
        moTable.TableRow(lRow) = Array(Join(moTable.TableRow(lRow), ""))
    Next
    
    Rerender
End Sub

Private Sub mnuTableSort_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    moTable.SortByColumn moCellPosition.ColumnIndex, False, False
    moSelectionn.SelectionType = stNone
    Rerender
End Sub

Private Sub mnuTableSortDescending_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    moTable.SortByColumn moCellPosition.ColumnIndex, True, False
    moSelectionn.SelectionType = stNone
    Rerender
End Sub

Private Sub mnuTableSortNumeric_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    moTable.SortByColumn moCellPosition.ColumnIndex, False, True
    moSelectionn.SelectionType = stNone
    Rerender
End Sub

Private Sub mnuTableSortNumericDescending_Click()
    moChangeHistory.Log moTable.Text, moSelectionn.EndPosition, moSelectionn.StartPosition
    
    moTable.SortByColumn moCellPosition.ColumnIndex, True, True
    moSelectionn.SelectionType = stNone
    Rerender
End Sub

Private Sub scrHorizontal_Change()
    moCursor.HideCursor
    goCanvas.Cls
    moRenderer.RenderTable
    moCursor.MoveCursor
End Sub

Private Sub scrHorizontal_Scroll()
    scrHorizontal_Change
End Sub

Private Sub scrVertical_Change()
    moCursor.HideCursor
    goCanvas.Cls
    moRenderer.RenderTable
    moCursor.MoveCursor
End Sub

Private Sub scrVertical_Scroll()
    scrVertical_Change
End Sub



Private Sub txtReplace_KeyPress(iKeyAscii As Integer)
    Dim sFind As String
    Dim sReplace As String
    Dim sText As String
    Dim sTableText As String
    Dim vColumns As Variant
    Dim lColumn As Long
    Dim lRow As Long
    
    If iKeyAscii = 13 Then
        sFind = moTable.DecodeString(txtFind.Text)
        sReplace = moTable.DecodeString(txtReplace.Text)
        
        Select Case moSelectionn.SelectionType
            Case stNone
                sTableText = moTable.Text
                moChangeHistory.Log sTableText, 0, 0
                moTable.Text = Replace$(sTableText, sFind, sReplace)
                moTable.ConvertTextToTable
                If Me.scrVertical.Value > moTable.LastRow Then
                    Me.scrVertical.Value = moTable.LastRow
                End If
                moCellPosition.TextPosition = 1
                Rerender
            Case stText
                sTableText = moTable.Text
                moChangeHistory.Log sTableText, 0, 0
                sText = Mid$(sTableText, moSelectionn.StartPosition + 1, moSelectionn.EndPosition - moSelectionn.StartPosition)
                sText = Replace$(sText, sFind, sReplace)
                moTable.Text = Left$(sTableText, moSelectionn.StartPosition) & sText & Mid$(sTableText, moSelectionn.EndPosition + 1)
                moSelectionn.EndPosition = moSelectionn.StartPosition + Len(sText)
                Rerender
            Case stColumns
                moChangeHistory.Log moTable.Text, 0, 0
                For lColumn = 0 To moTable.LastColumn
                    If moSelectionn.Column(lColumn) Then
                        For lRow = 0 To moTable.LastRow
                            sText = moTable.TableCell(lRow, lColumn)
                            sText = Replace$(sText, sFind, sReplace)
                            moTable.TableCell(lRow, lColumn) = sText
                        Next
                    End If
                Next
                Rerender
            Case stRows
                moChangeHistory.Log moTable.Text, 0, 0
                For lRow = 0 To moTable.LastRow
                    If moSelectionn.Row(lRow) Then
                        For lColumn = 0 To moTable.LastColumn
                            sText = moTable.TableCell(lRow, lColumn)
                            sText = Replace$(sText, sFind, sReplace)
                            moTable.TableCell(lRow, lColumn) = sText
                        Next
                    End If
                Next
                Rerender
            Case stCells
                moChangeHistory.Log moTable.Text, 0, 0
                For lColumn = 0 To moTable.LastColumn
                    For lRow = 0 To moTable.LastRow
                        If moSelectionn.Cell(lRow, lColumn) Then
                            sText = moTable.TableCell(lRow, lColumn)
                            sText = Replace$(sText, sFind, sReplace)
                            moTable.TableCell(lRow, lColumn) = sText
                        End If
                    Next
                Next
                Rerender
        End Select

        If scrVertical.Value > moTable.LastRow Then
            scrVertical.Value = moTable.LastRow
            moCursor.moPosition.RowIndex = moTable.LastRow
            Rerender
        End If

        txtReplace.SetFocus
    End If
End Sub

Private Sub txtReplace_LostFocus()
    txtReplace.TabStop = False
End Sub

Private Sub Rerender(Optional ByVal bClear As Boolean = True, Optional ByVal bUpdateCursor As Boolean = True)
    moCursor.HideCursor
    If bClear Then
        goCanvas.Cls
    End If
    moRenderer.RenderTable
    If bUpdateCursor Then
        UpdateCursor
    End If
End Sub

Private Sub UpdateCursor(Optional ByVal bRerender As Boolean = False)
    moCursor.MoveCursor
    If Not bRerender Then
        If moCursor.moPosition.RowIndex < scrVertical.Value Then
            scrVertical.Value = moCursor.moPosition.RowIndex
            Rerender
            If moSelectionn.SelectionType = stText Then
                moCursor.RecreateCursor
            End If
        End If
        If moCursor.moPosition.RowIndex > (Int(goCanvas.Height / moTableInfo.mlCellHeight) + scrVertical.Value - 2) Then
            scrVertical.Value = moCursor.moPosition.RowIndex - Int(goCanvas.Height / moTableInfo.mlCellHeight) + 2
            Rerender
            If moSelectionn.SelectionType = stText Then
                moCursor.RecreateCursor
            End If
        End If
        
        If moCursor.moPosition.ColumnIndex < scrHorizontal.Value Then
            scrHorizontal.Value = moCursor.moPosition.ColumnIndex
            Rerender
            If moSelectionn.SelectionType = stText Then
                moCursor.RecreateCursor
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtFind.Width = Me.Width - 1200
    txtReplace.Width = txtFind.Width
    goPanel.Width = Me.Width / Screen.TwipsPerPixelX
    goPanel.Top = Me.Height / Screen.TwipsPerPixelY - goPanel.Height - 50
    goCanvas.Width = Me.Width / Screen.TwipsPerPixelX - scrVertical.Width - 10
    goCanvas.Height = Me.Height / Screen.TwipsPerPixelX - goPanel.Height - 80
    scrVertical.Left = goCanvas.Width - 10
    scrVertical.Height = goCanvas.Height
    scrHorizontal.Top = goCanvas.Height + 2
    scrHorizontal.Width = goCanvas.Width
End Sub

Private Sub goCanvas_Paint()
    moRenderer.RenderTable
    moCursor.RecreateCursor
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbDefault
End Sub

Private Sub txtReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbDefault
End Sub
