VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Entries As New LetterEntry

Private Sub Form_Load()
    Decompose "HAMPSTEAD"
    Decompose "ROTHERHAM"
    Decompose "WAPPING"
    Decompose "ROTHERHITHE"
    Decompose "LITTLEHAMPTON"
    Decompose "BRIGHTON"
    List Entries, ""
End Sub

Private Sub AddPart(sPart As String)
    Dim lIndex As Long
    Dim SubEntry As LetterEntry
    Dim ThisEntry As LetterEntry
    
    Set ThisEntry = Entries
    
    ' Find string
    For lIndex = 1 To Len(sPart)
        Set ThisEntry = ThisEntry.Add(Mid$(sPart, lIndex, 1))
    Next
End Sub


Private Sub Decompose(sString As String)

    Dim lIndex As Long
    Dim lLength As Long
    
    For lIndex = 1 To Len(sString)
        AddPart Mid$(sString, lIndex)
    Next
End Sub

Private Sub List(oThisEntry As LetterEntry, ByVal sString As String)
    Dim lIndex As Long
    Dim dFinalScore As Double
    For lIndex = 0 To oThisEntry.Count - 1
        sString = sString & oThisEntry.Item(lIndex).Letter
        List oThisEntry.Item(lIndex), sString
        dFinalScore = Len(sString) ^ (oThisEntry.Item(lIndex).Score - 1)
        If dFinalScore > 1 Then
            Debug.Print sString, dFinalScore
        End If
        sString = ""
    Next
End Sub
