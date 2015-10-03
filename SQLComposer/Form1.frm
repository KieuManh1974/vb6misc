VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SQL Composer"
   ClientHeight    =   12015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
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
   ScaleHeight     =   12015
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDatabases 
      Height          =   1035
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox txtSQL 
      Height          =   3375
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.ListBox lstSQL 
      Height          =   2985
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox lstTables 
      Height          =   5325
      Left            =   0
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ListBox lstFields 
      Height          =   6495
      Left            =   3000
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const VK_LCONTROL = &HA2

Private sSelectedTable As String
Private oSelectedFields As New Dictionary
Private oSelectedTables As New Dictionary
Private oSelectedAliases As New Dictionary

Private Enum SelectionType
    stSelect
    stFilter
    stOrder
End Enum

Private mlSelectionType As SelectionType

Private Sub Form_Load()
    Set goDatabases.ListBoxControl = lstDatabases
    Set goTables.ListBoxControl = lstTables
    Set goFields.ListBoxControl = lstFields
    Set goSQL.ListBoxControl = lstSQL
    Set goSQL.TextBoxControl = txtSQL
    
    Set goCon = New Connection
    Set goConEBS = New Connection
    goCon.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & App.Path & "\joins.mdb" & ";Uid=admin;Pwd="
    goConEBS.Open "Provider=msdaora.1;Data Source=students;User id=fes;Password=buttercup;"

    InitialiseParser

    goDatabases.PopulateDatabases
    goTables.PopulateTables
    goExpressions.PopulateExpressions
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lstTables.Height = ScaleHeight
    lstFields.Height = ScaleHeight
    lstSQL.Height = ScaleHeight / 2
    lstSQL.Width = ScaleWidth - lstSQL.Left
    txtSQL.Height = ScaleHeight / 2
    txtSQL.Width = lstSQL.Width
    txtSQL.Top = lstSQL.Top + lstSQL.Height + 30
End Sub

Private Sub InitialiseParser()
    Dim sDef As String
    Dim oFSO As New FileSystemObject
    
    sDef = oFSO.OpenTextFile(App.Path & "/SQLComposer.pdl").ReadAll
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        Stop
    End If
    
    Set goParseExpression = ParserObjects("expression")
    Set goParseFunction = ParserObjects("function")
    Set goParseIdentifier = ParserObjects("identifier")
End Sub


Private Sub lstDatabases_Click()
    Dim oSelectedDatabase As clsDatabaseInfo
    
    Set oSelectedDatabase = goDatabases.GetDatabaseByIdentifier(lstDatabases.ItemData(lstDatabases.ListIndex))
    Set goTables.SelectedDatabase = oSelectedDatabase
    Set goFields.SelectedTable = Nothing
    Set goFields.SelectedExpressions = Nothing
    Set goFields.SelectedFields = Nothing
    Set goFields.SelectedFilters = Nothing
    
    goSQL.Display
End Sub

Private Sub lstSQL_Click()
    Dim lIndex As Long
    Dim sClipboard As String
    
    For lIndex = 0 To lstSQL.ListCount - 1
        sClipboard = sClipboard & lstSQL.List(lIndex) & vbCrLf
    Next
    Clipboard.SetText sClipboard
End Sub

Private Sub lstTables_Click()
    Dim oSelectedTable As clsTableInfo
    Dim bLeftControl As Boolean
    Static bIgnore As Boolean
    
    If bIgnore Then
        Exit Sub
    End If
    
    bLeftControl = GetKeyState(VK_LCONTROL) < 0
    
    Set oSelectedTable = goTables.GetTableByIdentifier(lstTables.ItemData(lstTables.ListIndex))
    If Not bLeftControl Then
        goTables.ToggleSelection oSelectedTable
    Else
        bIgnore = True
        lstTables.Selected(lstTables.ListIndex) = Not lstTables.Selected(lstTables.ListIndex)
        bIgnore = False
    End If
    Set goFields.SelectedTable = oSelectedTable
    goSQL.Display
End Sub


Private Sub lstFields_Click()
    Select Case goFields.GetLineByIdentifier(lstFields.ItemData(lstFields.ListIndex)).LineType
        Case fltExpression
            goFields.ToggleExpression goFields.GetLineByIdentifier(lstFields.ItemData(lstFields.ListIndex)).Expression
        Case fltField
            Select Case mlSelectionType
                Case stSelect
                    goFields.ToggleSelection goFields.GetLineByIdentifier(lstFields.ItemData(lstFields.ListIndex)).Field
                Case stFilter
                    goFields.ToggleFilter goFields.GetLineByIdentifier(lstFields.ItemData(lstFields.ListIndex)).Field
                Case stOrder
            End Select
    End Select
    
    goSQL.Display
End Sub

Private Sub lstFields_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Shift
        Case 0, 3
            mlSelectionType = stSelect
        Case 1
            mlSelectionType = stFilter
        Case 2
            mlSelectionType = stOrder
    End Select
End Sub

