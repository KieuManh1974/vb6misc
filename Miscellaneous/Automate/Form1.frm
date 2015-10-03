VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Automate"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3870
      Width           =   1455
   End
   Begin VB.ListBox lstScripts 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRun_Click()
    Dim oTree As New ParseTree
    
    Stream.Text = LoadScript(lstScripts.Text)
    If oParser.Parse(oTree) Then
        ExecuteScript oTree
    End If
End Sub

Private Sub Form_Load()
    Dim oTree As New ParseTree
    
    PopulateScripts
    Definition.Initialise
End Sub

Private Sub Form_Resize()
    lstScripts.Width = Me.ScaleWidth
    lstScripts.Height = cmdRun.Top - 50
End Sub

Private Sub PopulateScripts()
    Dim sExt As String
    Dim iDot As Integer
    Dim oFile As File
    
    With New FileSystemObject
        For Each oFile In .GetFolder(App.Path).Files
            iDot = InStrRev(oFile.Name, ".")
            
            If iDot > 0 Then
                sExt = UCase$(Mid$(oFile.Name, iDot + 1))
                If sExt = "AUT" Then
                    lstScripts.AddItem oFile.Name
                End If
            End If
    
        Next
    End With
End Sub
