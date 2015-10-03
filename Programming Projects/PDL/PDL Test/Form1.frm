VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oParseTest As IParseObject

Private Sub Form_Load()
    Dim sDef As String
    Dim oResult As New ParseTree

    sDef = sDef & "test := PERM 'a', 'b', 'c'; "

    If Not SetNewDefinition(sDef) Then
        MsgBox "BAD DEF"
    End If
    
    Set oParseTest = ParserObjects("test")

    Stream.Text = "xyz"
    If oParseTest.Parse(oResult) Then
    
    End If
End Sub
