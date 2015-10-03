VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PositionsRef As PositionList
Public RelationsRef As RelationshipList

Public FileStore As String
Public DiagramRef As Diagram

Public Sub ReadFile()
    Dim oPosition As Position
    Dim oRelationship As Relationship
    Dim oFile As TextStream
    Dim sLine As String
    Dim oFSO As New FileSystemObject
    Dim oTree As ParseTree
    
    If oFSO.FileExists(FileStore) Then
        Set oFile = oFSO.OpenTextFile(FileStore, ForReading)
        While Not oFile.AtEndOfStream
            sLine = oFile.ReadLine
            
            Set oTree = New ParseTree
            Stream.Text = sLine
            If oParseLine.Parse(oTree) Then
                Select Case oTree.Index
                    Case 1 ' Position
                        Set oPosition = New Position
                        With oPosition
                            .Reference = oTree(1)(1).Text
                            .Name = oTree(1)(2).Text
                            .PosX = oTree(1)(3).Text
                            .PosY = oTree(1)(4).Text
                            Set .DiagramRef = DiagramRef
                        End With
                        DiagramRef.Positions.List.Add oPosition
                    Case 2 ' Relation
                        Set oRelationship = New Relationship
                        With oRelationship
                            Set .FromPos = DiagramRef.Positions.FindPositionFromReference(oTree(1)(1).Text)
                            Set .ToPos = DiagramRef.Positions.FindPositionFromReference(oTree(1)(2).Text)
                            Set .DiagramRef = DiagramRef
                        End With
                        DiagramRef.Relations.List.Add oRelationship
                    Case 3 ' Offset
                        DiagramRef.sTopLeftX = oTree(1)(1).Text
                        DiagramRef.sTopLeftY = oTree(1)(2).Text
                    Case 4 ' Zoom
                End Select
            End If
            
        Wend
    End If
End Sub

Public Sub WriteFile()
    Dim oPosition As Position
    Dim oRelationship As Relationship
    Dim oFSO As New FileSystemObject
    Dim oFile As TextStream
    
    Set oFile = oFSO.OpenTextFile(FileStore, ForWriting, True)
    oFile.WriteLine "O:" & DiagramRef.sTopLeftX & "|" & DiagramRef.sTopLeftY
    For Each oPosition In PositionsRef.List
        With oPosition
            oFile.WriteLine "P:" & .Reference & "|" & .Name & "|" & .PosX & "|" & .PosY
        End With
    Next
    
    For Each oRelationship In RelationsRef.List
        With oRelationship
            oFile.WriteLine "R:" & .FromPos.Reference & "|" & .ToPos.Reference
        End With
    Next
    oFile.Close
End Sub


