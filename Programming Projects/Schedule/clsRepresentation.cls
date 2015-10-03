VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Displaylist As clsDisplayList
Public List As clsList
Public Tree As clsTree

Public Sub LoadNotes()
    Dim oFSO As New FileSystemObject
    
    If oFSO.FileExists(App.Path & "\tasks.txt") Then
        Notes = oFSO.OpenTextFile(App.Path & "\tasks.txt", ForReading, , TristateTrue).ReadAll
    Else
        Notes = ""
    End If
End Sub

Public Sub SaveNotes()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    
    oFSO.CreateTextFile(App.Path & "\tasks.txt", True, True).Write Notes
End Sub

Public Property Let Notes(sNotes As String)
    Set List = New clsList
    List.Notes = sNotes
    RefreshDisplayListFromList
    RefreshTreeFromList
End Property

Public Property Get Notes() As String
    Notes = List.Notes
End Property

Private Sub RefreshTreeFromList()
    Dim vNotes As Variant
    Dim oHierarchy As clsHierarchy
    Dim oStack As New Collection
    Dim sNote As String
    Dim lLevel As Long
    Dim lPreviousLevel As Long
    Dim oContext As clsHierarchy
    Dim oPreviousHierarchy As clsHierarchy
    Dim lPopIndex As Long
    
    Set Tree = New clsTree
    
    Set Tree.Top = New clsHierarchy
    oStack.Add Tree.Top
    Set oContext = Tree.Top
        
    lPreviousLevel = 0
    For Each oHierarchy In List.List
        lLevel = oHierarchy.Level
        If lLevel > lPreviousLevel Then
            oStack.Add oPreviousHierarchy
            Set oContext = oPreviousHierarchy
        ElseIf lLevel < lPreviousLevel Then
            For lPopIndex = lLevel To lPreviousLevel - 1
                oStack.Remove (oStack.Count)
            Next
            Set oContext = oStack.Item(oStack.Count)
        End If
        
        Set oHierarchy.Parent = oContext
        oContext.Children.Add oHierarchy
        
        Set oPreviousHierarchy = oHierarchy
        lPreviousLevel = lLevel
    Next
End Sub

Private Sub RefreshListFromTree()
    Set List = New clsList
    RefreshListFromChildren Tree.Top
    List.List.Remove (1)
End Sub

Private Sub RefreshListFromChildren(oHierarchy As clsHierarchy)
    Dim oChild As clsHierarchy
    
    List.List.Add oHierarchy
    For Each oChild In oHierarchy.Children
        RefreshListFromChildren oChild
    Next
End Sub

Public Sub RefreshDisplayListFromList()
    Dim oHierarchy As clsHierarchy
    Dim bShow As Boolean
    Dim lLevel As Long
    
    Set Displaylist = New clsDisplayList
    bShow = True
    
    For Each oHierarchy In List.List
        If oHierarchy.Level <= lLevel Then
            bShow = True
        End If
        If bShow Then
            Displaylist.List.Add oHierarchy
            If Not oHierarchy.Expanded Then
                bShow = False
                lLevel = oHierarchy.Level
            End If
        End If
    Next
End Sub

Public Function NewTask(oParent As clsHierarchy, sText As String) As Long
    Dim oNewHierarchy As New clsHierarchy
    
    oNewHierarchy.Task.Description = sText
    Set oNewHierarchy.Parent = oParent
    oNewHierarchy.Identifier = List.NewIdentifier
    oNewHierarchy.Level = oParent.Level + 1
    oParent.Children.Add oNewHierarchy
    RefreshListFromTree
    RefreshDisplayListFromList
    NewTask = Displaylist.FindIdentifierIndex(oParent.Identifier)
End Function

Public Function RemoveTask(oTask As clsHierarchy) As Long
    Dim oParent As clsHierarchy
    
    Set oParent = oTask.Parent
    oParent.Children.Remove (oParent.FindIdentifierIndex(oTask.Identifier))
    Set oTask = Nothing
    RefreshListFromTree
    RefreshDisplayListFromList
    RemoveTask = Displaylist.FindIdentifierIndex(oParent.Identifier)
End Function

Public Function MoveTaskUp(oTask As clsHierarchy) As Long
    Dim oParent As clsHierarchy
    Dim lChildIndex As Long
    
    Set oParent = oTask.Parent
    lChildIndex = oParent.FindIdentifierIndex(oTask.Identifier)
    
    If lChildIndex > 1 Then
        oParent.Children.Remove lChildIndex
        oParent.Children.Add oTask, , lChildIndex - 1
        RefreshListFromTree
        RefreshDisplayListFromList
        MoveTaskUp = Displaylist.FindIdentifierIndex(oTask.Identifier)
    Else
        MoveTaskUp = -1
    End If
End Function

Public Function MoveTaskDown(oTask As clsHierarchy) As Long
    Dim oParent As clsHierarchy
    Dim lChildIndex As Long
    
    Set oParent = oTask.Parent
    lChildIndex = oParent.FindIdentifierIndex(oTask.Identifier)
    
    If lChildIndex < oParent.Children.Count Then
        oParent.Children.Add oTask, , , lChildIndex + 1
        oParent.Children.Remove lChildIndex
        RefreshListFromTree
        RefreshDisplayListFromList
        MoveTaskDown = Displaylist.FindIdentifierIndex(oTask.Identifier)
    Else
        MoveTaskDown = -1
    End If
End Function

Public Function MoveTask(oTask As clsHierarchy, oNewPosition As clsHierarchy) As Long
    Dim oCopyTask As clsHierarchy
    
    Set oCopyTask = oTask
    
    RemoveTask oTask
    Set oCopyTask.Parent = oNewPosition
    oNewPosition.Children.Add oCopyTask
    oNewPosition.Expanded = True
    Tree.RefreshLevels
    RefreshListFromTree
    RefreshDisplayListFromList
    MoveTask = Displaylist.FindIdentifierIndex(oCopyTask.Identifier)
End Function
