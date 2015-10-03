VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Property Let Notes(sNotes As String)
    Dim lNoteIndex As Long
    Dim oHierarchy As clsHierarchy
    Dim vNotes As Variant
    Dim sNote As String
    
    vNotes = Split(sNotes, vbCrLf)
    For lNoteIndex = LBound(vNotes) To UBound(vNotes)
        sNote = vNotes(lNoteIndex)
        Set oHierarchy = New clsHierarchy
        oHierarchy.Descriptor = sNote
        oHierarchy.Identifier = lNoteIndex
        List.Add oHierarchy
    Next
End Property

Public Property Get Notes() As String
    Dim lNoteIndex As Long
    Dim oHierarchy As clsHierarchy
    
    For Each oHierarchy In List
        Notes = Notes & oHierarchy.Descriptor & vbCrLf
    Next
    If Notes <> "" Then
        Notes = Left$(Notes, Len(Notes) - 2)
    End If
End Property

Public Function NewIdentifier() As Long
    Dim oHierarchy As clsHierarchy
    
    NewIdentifier = -1
    For Each oHierarchy In List
        If oHierarchy.Identifier > NewIdentifier Then
            NewIdentifier = oHierarchy.Identifier
        End If
    Next
    NewIdentifier = NewIdentifier + 1
End Function
