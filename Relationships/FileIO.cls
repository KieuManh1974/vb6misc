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
Public CanvasRef As Form

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
            ParserTextString.ParserText = sLine
            If oParsePosition.Parse(oTree) Then
                Set oPosition = New Position
                With oPosition
                    .Reference = oTree(1).Text
                    .Name = oTree(2).Text
                    .PosX = oTree(3).Text
                    .PosY = oTree(4).Text
                    Set .CanvasRef = CanvasRef
                End With
                Positions.List.Add oPosition
            End If
            Set oTree = New ParseTree
            ParserTextString.ParserText = sLine
            If oParseRelationship.Parse(oTree) Then
                Set oRelationship = New Relationship
                With oRelationship
                    Set .FromPos = Positions.FindPositionFromReference(oTree(1).Text)
                    Set .ToPos = Positions.FindPositionFromReference(oTree(2).Text)
                    Set .CanvasRef = CanvasRef
                End With
                Relations.List.Add oRelationship
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


