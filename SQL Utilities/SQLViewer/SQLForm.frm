VERSION 5.00
Begin VB.Form SQLForm 
   Caption         =   "SQL Viewer"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   16125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   16125
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSplitter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   4680
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2445
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   0
      Width           =   40
   End
   Begin VB.TextBox txtSQLOutput 
      Height          =   2535
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.TextBox txtSQLInput 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
End
Attribute VB_Name = "SQLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum FocusTypes
    InputPane
    OutputPane
End Enum

Private bDragSplitter As Boolean
Private iFocusWindow As FocusTypes

Private Sub Form_Load()
    With oSQLFormatter
        .KeywordStyle = GetSetting("SQLViewer", "Options", "KeywordStyle", .KeywordStyle)
        .OperatorStyle = GetSetting("SQLViewer", "Options", "OperatorStyle", .OperatorStyle)
        .FunctionStyle = GetSetting("SQLViewer", "Options", "FunctionStyle", .FunctionStyle)
        .FieldStyle = GetSetting("SQLViewer", "Options", "FieldStyle", .FieldStyle)
        .TableStyle = GetSetting("SQLViewer", "Options", "TableStyle", .TableStyle)
        .QuoteStyle = GetSetting("SQLViewer", "Options", "QuoteStyle", .QuoteStyle)
        .WHEREExpression = GetSetting("SQLViewer", "Options", "WHEREExpression", .WHEREExpression)
        .VisualBasicCode = GetSetting("SQLViewer", "Options", "VisualBasicCode", .VisualBasicCode)
        .TableAliasing = GetSetting("SQLViewer", "Options", "TableAliasing", .TableAliasing)
        .ASSelectStyle = GetSetting("SQLViewer", "Options", "ASSelectStyle", .ASSelectStyle)
        .ASFromStyle = GetSetting("SQLViewer", "Options", "ASFromStyle", .ASFromStyle)
        .VBSQLIndex = GetSetting("SQLViewer", "Options", "VBSQLIndex", .VBSQLIndex)
        .Indented = GetSetting("SQLViewer", "Options", "Indented", .Indented)
        .Style = GetSetting("SQLViewer", "Options", "Style", .Style)
        .IndentedSubQuery = GetSetting("SQLViewer", "Options", "IndentedSubQuery", .IndentedSubQuery)
    End With
    
    Me.Width = GetSetting("SQLViewer", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("SQLViewer", "Dimensions", "Height", Me.Height)
    pctSplitter.Left = GetSetting("SQLViewer", "Dimensions", "Splitter", pctSplitter.Left)
    Me.WindowState = GetSetting("SQLViewer", "Dimensions", "WindowState", Me.WindowState)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pctSplitter.Height = Me.ScaleHeight
    txtSQLInput.Height = Me.ScaleHeight
    txtSQLOutput.Height = Me.ScaleHeight
    
    txtSQLInput.Width = pctSplitter.Left
    txtSQLOutput.Width = Me.ScaleWidth - pctSplitter.Left - pctSplitter.Width
    txtSQLOutput.Left = pctSplitter.Left + pctSplitter.Width
    
    If Me.WindowState = 0 Then
        SaveSetting "SQLViewer", "Dimensions", "Width", Me.Width
        SaveSetting "SQLViewer", "Dimensions", "Height", Me.Height
        SaveSetting "SQLViewer", "Dimensions", "Splitter", pctSplitter.Left
    End If
End Sub

Private Sub Form_Terminate()
    With oSQLFormatter
        SaveSetting "SQLViewer", "Options", "KeywordStyle", .KeywordStyle
        SaveSetting "SQLViewer", "Options", "OperatorStyle", .OperatorStyle
        SaveSetting "SQLViewer", "Options", "FunctionStyle", .FunctionStyle
        SaveSetting "SQLViewer", "Options", "FieldStyle", .FieldStyle
        SaveSetting "SQLViewer", "Options", "TableStyle", .TableStyle
        SaveSetting "SQLViewer", "Options", "QuoteStyle", .QuoteStyle
        SaveSetting "SQLViewer", "Options", "WHEREExpression", .WHEREExpression
        SaveSetting "SQLViewer", "Options", "VisualBasicCode", .VisualBasicCode
        SaveSetting "SQLViewer", "Options", "TableAliasing", .TableAliasing
        SaveSetting "SQLViewer", "Options", "ASSelectStyle", .ASSelectStyle
        SaveSetting "SQLViewer", "Options", "ASFromStyle", .ASFromStyle
        SaveSetting "SQLViewer", "Options", "VBSQLIndex", .VBSQLIndex
        SaveSetting "SQLViewer", "Options", "Indented", .Indented
        SaveSetting "SQLViewer", "Options", "Style", .Style
        SaveSetting "SQLViewer", "Options", "IndentedSubQuery", .IndentedSubQuery
                
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
     SaveSetting "SQLViewer", "Dimensions", "WindowState", Me.WindowState
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOptions_Click()
    Dim bOK As Boolean
    Dim lErrorPos As Long
    
    OptionsChanged = False
    ViewerOptions.Show vbModal
    
    If OptionsChanged Then
        If iFocusWindow = InputPane Then
            txtSQLOutput.Text = oSQLFormatter.FormatSQL(txtSQLInput.Text, bOK, lErrorPos)
        Else
            txtSQLOutput.Text = oSQLFormatter.FormatSQL(txtSQLOutput.Text, bOK, lErrorPos)
        End If
        If Not bOK Then
            txtSQLInput.ForeColor = vbRed
            txtSQLInput.SelStart = lErrorPos
            txtSQLInput.SelLength = 0
        Else
            txtSQLInput.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub pctSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        pctSplitter.Left = pctSplitter.Left + X
    End If
End Sub

Private Sub pctSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Form_Resize
    End If
End Sub

Private Sub txtSQLInput_Change()
    txtSQLInput.ForeColor = vbBlack
End Sub


Private Sub txtSQLInput_GotFocus()
    iFocusWindow = InputPane
End Sub

Private Sub txtSQLOutput_GotFocus()
    iFocusWindow = OutputPane
End Sub


Private Sub txtSQLInput_KeyPress(KeyAscii As Integer)
    Dim bOK As Boolean
    Dim lErrorPos As Long
    
    Select Case KeyAscii
        Case 1
            txtSQLInput.SelStart = 0
            txtSQLInput.SelLength = Len(txtSQLInput.Text)
        Case 10
            KeyAscii = 0
            Me.MousePointer = vbHourglass
            txtSQLOutput.Text = oSQLFormatter.FormatSQL(txtSQLInput.Text, bOK, lErrorPos)
            Me.MousePointer = vbDefault
            If Not bOK Then
                txtSQLInput.ForeColor = vbRed
                txtSQLInput.SelStart = lErrorPos
                txtSQLInput.SelLength = 0
            Else
                txtSQLInput.ForeColor = vbBlack
                txtSQLOutput.ForeColor = vbBlack
            End If
        Case 27
            txtSQLInput.Text = ""
    End Select
End Sub

Private Sub txtSQLOutput_KeyPress(KeyAscii As Integer)
    Dim bOK As Boolean
    Dim lErrorPos As Long
    
    Select Case KeyAscii
        Case 1
            txtSQLOutput.SelStart = 0
            txtSQLOutput.SelLength = Len(txtSQLOutput.Text)
        Case 10
            KeyAscii = 0
            Me.MousePointer = vbHourglass
            txtSQLOutput.Text = oSQLFormatter.FormatSQL(txtSQLOutput.Text, bOK, lErrorPos)
            Me.MousePointer = vbDefault
            If Not bOK Then
                txtSQLOutput.ForeColor = vbRed
                txtSQLOutput.SelStart = lErrorPos
                txtSQLOutput.SelLength = 0
            Else
                txtSQLOutput.ForeColor = vbBlack
            End If
        Case 27
            txtSQLOutput.Text = ""
    End Select
End Sub
