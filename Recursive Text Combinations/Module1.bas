Attribute VB_Name = "Structure"
Option Explicit

Public Function CreateStructure(oTree As ParseTree) As String
    Dim oEntry As Entry

    Set oEntry = CreateList(oTree)
'    CreateStructure = oEntry.EnumerateSequence
    
    oEntry.EnumerateList
    CreateStructure = oEntry.FetchSetText
End Function

Private Function CreateList(oTree As ParseTree) As Entry
    Dim oEntry As Entry
    Dim oListItem As ParseTree
    Dim oItem As ParseTree
    Dim sIndexName As String
    Dim oSubEntry As Entry
    Dim oIndexRef As Entry
    
    Set CreateList = New Entry
    Set CreateList.EntryList = New Collection
    CreateList.EntryType = etList
    
    If oTree(1).Index = 1 Then
        CreateList.Invisible = True
    End If
    If oTree(2).Index = 1 Then
        CreateList.IndexName = oTree(2)(1).Text
    End If
    For Each oItem In oTree(3).SubTree
        CreateList.EntryList.Add CreateSequence(oItem)
    Next
End Function

Private Function CreateSequence(oTree As ParseTree) As Entry
    Dim oEntry As Entry
    Dim oListItem As ParseTree
    Dim oItem As ParseTree
    Dim sIndexName As String
    Dim oSubEntry As Entry
    Dim oIndexRef As Entry
    
    Set CreateSequence = New Entry
    Set CreateSequence.EntryList = New Collection
    CreateSequence.EntryType = etSequence
    
    For Each oItem In oTree.SubTree
        Select Case oItem.Index
            Case 1 'index
                Set oSubEntry = New Entry
                oSubEntry.EntryType = etIndex
                oSubEntry.IndexName = oItem(1).Text
                CreateSequence.EntryList.Add oSubEntry
            Case 2 'sub
                CreateSequence.EntryList.Add CreateList(oItem(1)(1))
            Case 3 'text
                Set oSubEntry = New Entry
                oSubEntry.EntryType = etText
                oSubEntry.EntryText = oItem(1).Text
                CreateSequence.EntryList.Add oSubEntry
        End Select
    Next
End Function

