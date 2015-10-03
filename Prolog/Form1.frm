VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtProgram 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'    Dim result As New Collection
'
'    Dim Expression As New CParseObject
'
'    Dim Char As New CParseObject
'    Char.CharacterRange False, Chr(32), Chr(255)
'
'    Dim OrSep As New CParseObject
'    OrSep.Literal True, "|"
'
'    Dim LeftBracket As New CParseObject
'    LeftBracket.Literal True, "["
'
'    Dim RightBracket As New CParseObject
'    RightBracket.Literal True, "]"
'
'    Dim EndOfLine As New CParseObject
'    EndOfLine.Literal True, Chr(0)
'
'    Dim SpecialChar As New CParseObject
'    SpecialChar.Choice False, False, False, OrSep, LeftBracket, RightBracket, EndOfLine
'
'    Dim NotSpecialChar As New CParseObject
'    NotSpecialChar.Neg SpecialChar
'
'    Dim LiteralString As New CParseObject
'    LiteralString.RepeatUntil False, False, SpecialChar, Char
'    LiteralString.ErrorString = "Literal"
'
'    Dim ChoiceStringRepeat As New CParseObject
'    ChoiceStringRepeat.RepeatUntil False, True, RightBracket, OrSep, LiteralString
'
'    Dim ChoiceString As New CParseObject
'    ChoiceString.Join False, True, LeftBracket, LiteralString, ChoiceStringRepeat, RightBracket
'
'    Dim Element As New CParseObject
'    Element.Choice False, True, False, ChoiceString, LiteralString
'
'    Dim SearchProgram As New CParseObject
'    SearchProgram.RepeatUntil False, True, EndOfLine, Element
'
'    With New CParseText
'        .ParseString = txtProgram.Text
'        Set result = .IParseObject_FindString(SearchProgram)
'   End With
'
'End Sub
'
Private Sub txtProgram_Change()

End Sub
