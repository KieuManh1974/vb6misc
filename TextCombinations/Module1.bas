Attribute VB_Name = "Module1"
Option Explicit

Public Const NOT_FOUND As Long = -1

Public Type Index
    Name As String
    Value As Long
    Max As Long
    AssociatedList As Long
End Type

Public Type Entry
    IndexPtr As Long
    StringPtr As Long
    EntryPtr As Long
    Invisible As Boolean
End Type

Public Type StringList
    List() As String
    ListNotEmpty As Boolean
End Type

Public Lists() As StringList
Public Indeces() As Index
Public Entries() As Entry

Private bListsNotEmpty As Boolean
Private bIndecesNotEmpty As Boolean
Private bEntriesNotEmpty As Boolean

Public oParser As IParseObject

Public Sub InitialiseParser()
    Dim sDefinition As String
    
    sDefinition = sDefinition & "text := {REPEAT OR '[[', '{{', IN 0 TO 255, NOT '[{'};"
    sDefinition = sDefinition & "inner_text := {REPEAT OR '[]', ']]', '}}', '||', IN 32 TO 255, NOT ']|}'};"
    sDefinition = sDefinition & "variable := {AND (IN CASE 'A' TO 'Z'), {REPEAT IN CASE 'A' TO 'Z', '0' TO '9' MIN 0}};"
    sDefinition = sDefinition & "def:= AND variable, [':'];"
    sDefinition = sDefinition & "list_visible :=  AND ['['], ?def, (LIST inner_text, ['|'] MIN 0), [']'];"
    sDefinition = sDefinition & "list_invisible :=  AND ['{'], ?def, (LIST inner_text, ['|'] MIN 0), ['}'];"
    sDefinition = sDefinition & "list:=  OR list_visible, list_invisible;"
    sDefinition = sDefinition & "multiplier := REPEAT OR text, list UNTIL EOS;"

    If Not SetNewDefinition(sDefinition) Then
        MsgBox "Def Error"
        End
    End If
    
    Set oParser = ParserObjects("multiplier")
End Sub

Public Function AddIndex(sName As String, iMax As Long, iAssociatedList As Long) As Long
    Dim tIndex As Index
    
    If bIndecesNotEmpty Then
        ReDim Preserve Indeces(UBound(Indeces) + 1) As Index
    Else
        ReDim Indeces(0) As Index
        bIndecesNotEmpty = True
    End If
    
    tIndex.Name = sName
    tIndex.Value = 0
    tIndex.Max = iMax
    tIndex.AssociatedList = iAssociatedList
    
    Indeces(UBound(Indeces)) = tIndex
    AddIndex = UBound(Indeces)
End Function

Public Function AddStringList(aList As StringList) As Long
    If bListsNotEmpty Then
        ReDim Preserve Lists(UBound(Lists) + 1) As StringList
    Else
        ReDim Lists(0) As StringList
        bListsNotEmpty = True
    End If

    Lists(UBound(Lists)) = aList
    AddStringList = UBound(Lists)
End Function

Public Function AddList(sString As String, aList As StringList) As Long
    If aList.ListNotEmpty Then
        ReDim Preserve aList.List(UBound(aList.List) + 1) As String
    Else
        ReDim aList.List(0) As String
        aList.ListNotEmpty = True
    End If

    aList.List(UBound(aList.List)) = sString
End Function

Public Sub AddEntry(iIndexPtr As Long, iListPtr As Long, bInvisible As Boolean)
    Dim tEntry As Entry
    
    If bEntriesNotEmpty Then
        ReDim Preserve Entries(UBound(Entries) + 1) As Entry
    Else
        ReDim Entries(0) As Entry
        bEntriesNotEmpty = True
    End If
    
    tEntry.IndexPtr = iIndexPtr
    tEntry.StringPtr = iListPtr
    tEntry.Invisible = bInvisible
    
    Entries(UBound(Entries)) = tEntry
End Sub

Public Function FindIndexByName(sVariable As String) As Long
    Dim iIndecesIndex As Long
    
    FindIndexByName = NOT_FOUND
    
    If sVariable = "" Then
        Exit Function
    End If
    If bIndecesNotEmpty Then
        For iIndecesIndex = 0 To UBound(Indeces)
            If Indeces(iIndecesIndex).Name = sVariable Then
                FindIndexByName = iIndecesIndex
                Exit Function
            End If
        Next
    End If
End Function

Public Function CreateStructure(oParseTree As ParseTree) As String
    Dim oPart As ParseTree
    Dim aStringList As StringList
    Dim iListPtr As Long
    Dim sIndex As String
    Dim iIndexPtr As Long
    Dim oChoice As ParseTree
    Dim bInvisible As Boolean
    
    Erase Entries
    Erase Indeces
    Erase Lists
    bListsNotEmpty = False
    bIndecesNotEmpty = False
    bEntriesNotEmpty = False
    
    For Each oPart In oParseTree.SubTree
        Select Case oPart.Index
            Case 1
                Erase aStringList.List
                aStringList.ListNotEmpty = False
                AddList oPart.Text, aStringList
                iListPtr = AddStringList(aStringList)
                AddEntry NOT_FOUND, iListPtr, False
            Case 2
                bInvisible = oPart(1).Index = 2
                sIndex = oPart(1)(1)(1).Text
                iIndexPtr = FindIndexByName(sIndex)
                If oPart(1)(1)(2).SubTree.Count > 0 Then
                    Erase aStringList.List
                    aStringList.ListNotEmpty = False
                    For Each oChoice In oPart(1)(1)(2).SubTree
                        If oChoice.Text = "[]" Then
                            AddList "", aStringList
                        Else
                            AddList oChoice.Text, aStringList
                        End If
                    Next
                    iListPtr = AddStringList(aStringList)
                    If iIndexPtr = NOT_FOUND Then
                        iIndexPtr = AddIndex(sIndex, oPart(1)(1)(2).SubTree.Count, iListPtr)
                    End If
                Else
                    If iIndexPtr = NOT_FOUND Then
                        'AddIndex sIndex, oPart(1)(2).SubTree.Count, iListPtr
                    Else
                        iListPtr = Indeces(iIndexPtr).AssociatedList
                    End If
                End If
                
                AddEntry iIndexPtr, iListPtr, bInvisible
        End Select
    Next
    
    CreateStructure = EnumerateStructure
End Function

Public Function EnumerateStructure() As String
    Dim vEntry As Entry
    Dim iEntryIndex As Long
    
    Do
        For iEntryIndex = 0 To UBound(Entries)
            vEntry = Entries(iEntryIndex)
            If Not vEntry.Invisible Then
                If vEntry.IndexPtr <> NOT_FOUND Then
                    EnumerateStructure = EnumerateStructure & Lists(vEntry.StringPtr).List(Indeces(vEntry.IndexPtr).Value)
                Else
                    EnumerateStructure = EnumerateStructure & Lists(vEntry.StringPtr).List(0)
                End If
            End If
        Next
        EnumerateStructure = EnumerateStructure & vbCrLf
    Loop Until EnumerateIndeces
End Function

Public Function EnumerateIndeces() As Boolean
    Dim iIndexPtr As Long
    Dim bExit As Boolean
    Dim iMaxIndexPtr As Long
    
    If Not bIndecesNotEmpty Then
        EnumerateIndeces = True
        Exit Function
    End If
    
    iMaxIndexPtr = UBound(Indeces)
    iIndexPtr = iMaxIndexPtr
    
    Do
        Indeces(iIndexPtr).Value = Indeces(iIndexPtr).Value + 1
        If Indeces(iIndexPtr).Value = Indeces(iIndexPtr).Max Then
            Indeces(iIndexPtr).Value = 0
            iIndexPtr = iIndexPtr - 1
            bExit = False
        Else
            bExit = True
        End If
    Loop Until bExit Or iIndexPtr < 0
    
    If iIndexPtr < 0 Then
        EnumerateIndeces = True
    End If
End Function
