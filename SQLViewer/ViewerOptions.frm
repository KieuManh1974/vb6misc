VERSION 5.00
Begin VB.Form ViewerOptions 
   Caption         =   "SQL Viewer - Options"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame11 
      Caption         =   "Indented Sub Query"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   60
      Top             =   1080
      Width           =   3975
      Begin VB.OptionButton optIndentedSubQuery 
         Caption         =   "Yes"
         Height          =   255
         Index           =   1
         Left            =   1095
         TabIndex        =   62
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optIndentedSubQuery 
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Functions"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   55
      Top             =   4440
      Width           =   3975
      Begin VB.OptionButton optFunctions 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   59
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optFunctions 
         Caption         =   "Proper"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   58
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optFunctions 
         Caption         =   "Lower"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   57
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optFunctions 
         Caption         =   "Upper"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Operators"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   50
      Top             =   3960
      Width           =   3975
      Begin VB.OptionButton optOperators 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   54
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optOperators 
         Caption         =   "Proper"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   53
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optOperators 
         Caption         =   "Lower"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   52
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optOperators 
         Caption         =   "Upper"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Field Space Format"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   47
      Top             =   6840
      Width           =   3975
      Begin VB.OptionButton optStyle 
         Caption         =   "Access"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton optStyle 
         Caption         =   "MySQL"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   48
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Indented"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   600
      Width           =   3975
      Begin VB.OptionButton optIndented 
         Caption         =   "Yes"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   46
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optIndented 
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "VB SQL Index"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   1560
      Width           =   3975
      Begin VB.OptionButton optVBSQLIndex 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   43
         Top             =   180
         Width           =   495
      End
      Begin VB.OptionButton optVBSQLIndex 
         Caption         =   "2"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   42
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optVBSQLIndex 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton optVBSQLIndex 
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   40
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "AS (FROM)"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   3975
      Begin VB.OptionButton optASFrom 
         Caption         =   "Leave"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   38
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optASFrom 
         Caption         =   "AS"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   37
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton optASFrom 
         Caption         =   "Blank"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame13 
      Caption         =   "Table Aliases"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   3975
      Begin VB.OptionButton optTableAliases 
         Caption         =   "Un-alias"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optTableAliases 
         Caption         =   "No Alias"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   32
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optTableAliases 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   31
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTableAliases 
         Caption         =   "Alias"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Visual Basic Code"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optVisualBasicCode 
         Caption         =   "VB"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton optVisualBasicCode 
         Caption         =   "SQL"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   27
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "WHERE Expression"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   3975
      Begin VB.OptionButton optWhereExpression 
         Caption         =   "Leave"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optWhereExpression 
         Caption         =   "Indent"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "AS (SELECT)"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   3975
      Begin VB.OptionButton optASSelect 
         Caption         =   "Blank"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton optASSelect 
         Caption         =   "AS"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   21
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton optASSelect 
         Caption         =   "Leave"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   20
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Quotes"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   3975
      Begin VB.OptionButton optQuotes 
         Caption         =   "Leave"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   18
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optQuotes 
         Caption         =   "Double"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton optQuotes 
         Caption         =   "Single"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Tables"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   3975
      Begin VB.OptionButton optTables 
         Caption         =   "Upper"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optTables 
         Caption         =   "Lower"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optTables 
         Caption         =   "Proper"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optTables 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Fields"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
      Begin VB.OptionButton optFields 
         Caption         =   "Upper"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optFields 
         Caption         =   "Lower"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optFields 
         Caption         =   "Proper"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optFields 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keywords"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   3975
      Begin VB.OptionButton optKeywords 
         Caption         =   "Leave"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optKeywords 
         Caption         =   "Proper"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optKeywords 
         Caption         =   "Lower"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optKeywords 
         Caption         =   "Upper"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   975
      End
   End
End
Attribute VB_Name = "ViewerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    OptionsChanged = True
    Unload Me
End Sub

Private Sub Form_Load()
    PopulateOptions
End Sub

Private Sub PopulateOptions()
    With oSQLFormatter
        optTables(.TableStyle).Value = True
        optOperators(.OperatorStyle).Value = True
        optKeywords(.KeywordStyle).Value = True
        optFunctions(.FunctionStyle).Value = True
        optFields(.FieldStyle).Value = True
        optQuotes(.QuoteStyle).Value = True
        optASSelect(.ASSelectStyle).Value = True
        optASFrom(.ASFromStyle).Value = True
        optWhereExpression(.WHEREExpression).Value = True
        optVisualBasicCode(.VisualBasicCode).Value = True
        optTableAliases(.TableAliasing).Value = True
        optVBSQLIndex(.VBSQLIndex) = True
        optIndented(.Indented) = True
        optStyle(.Style) = True
        optIndentedSubQuery(.IndentedSubQuery) = True
    End With
End Sub

Private Sub optASFrom_Click(Index As Integer)
    oSQLFormatter.ASFromStyle = Index
End Sub

Private Sub optASSelect_Click(Index As Integer)
    oSQLFormatter.ASSelectStyle = Index
End Sub

Private Sub optFields_Click(Index As Integer)
    oSQLFormatter.FieldStyle = Index
End Sub

Private Sub optFunctions_Click(Index As Integer)
    oSQLFormatter.FunctionStyle = Index
End Sub

Private Sub optIndented_Click(Index As Integer)
    oSQLFormatter.Indented = Index
End Sub

Private Sub optIndentedSubQuery_Click(Index As Integer)
    oSQLFormatter.IndentedSubQuery = Index
End Sub

Private Sub optKeywords_Click(Index As Integer)
    oSQLFormatter.KeywordStyle = Index
End Sub

Private Sub optOperators_Click(Index As Integer)
    oSQLFormatter.OperatorStyle = Index
End Sub

Private Sub optQuotes_Click(Index As Integer)
    oSQLFormatter.QuoteStyle = Index
End Sub

Private Sub optStyle_Click(Index As Integer)
    oSQLFormatter.Style = Index
End Sub

Private Sub optTableAliases_Click(Index As Integer)
    oSQLFormatter.TableAliasing = Index
End Sub

Private Sub optTables_Click(Index As Integer)
    oSQLFormatter.TableStyle = Index
End Sub

Private Sub optVisualBasicCode_Click(Index As Integer)
    oSQLFormatter.VisualBasicCode = Index
End Sub

Private Sub optWhereExpression_Click(Index As Integer)
    oSQLFormatter.WHEREExpression = Index
End Sub

Private Sub optVBSQLIndex_Click(Index As Integer)
    oSQLFormatter.VBSQLIndex = Index
End Sub

