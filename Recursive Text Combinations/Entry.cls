VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum etEntryTypes
    etSequence
    etList
    etIndex
    etText
End Enum

Public EntryType As etEntryTypes
Public IndexPtr As Entry
Public EntryText As String
Public EntryList As Collection
Public SetList As Collection

Public IndexName As String
Public IndexValue As Long
Public IndexMax As Long

Public Invisible As Boolean

Private Sub Class_Initialize()
    IndexValue = 1
End Sub

Public Function EnumerateSequence() As String
    Dim iIndexPtr As Long
    Dim bExit As Boolean
    Dim iMaxIndexPtr As Long
    Dim oSubEntry As Entry
    Dim oIndexEntry As Entry
    Dim sEnumerateText As String
    Dim oSubEntries As Collection
    Dim sEnumeratedSequence As String
    
    On Error Resume Next
    
    Set oSubEntries = Me.EntryList
    
    If oSubEntries.Count = 0 Then
        Exit Function
    End If

    iMaxIndexPtr = oSubEntries.Count
    Do
        For Each oSubEntry In oSubEntries
            Select Case oSubEntry.EntryType
                Case etSequence
                    sEnumeratedSequence = oSubEntry.EnumerateSequence
                    EnumerateSequence = EnumerateSequence & sEnumeratedSequence
                Case etList
                    If Not oSubEntry.Invisible Then
                        Set oIndexEntry = FindIndex(oSubEntry.IndexName, oSubEntries)
                        If oSubEntry Is oIndexEntry Or oIndexEntry Is Nothing Then
                            sEnumeratedSequence = oSubEntry.EntryList(oSubEntry.IndexValue).EnumerateSequence
                            EnumerateSequence = EnumerateSequence & sEnumeratedSequence
                        Else
                            sEnumeratedSequence = oSubEntry.EntryList(oIndexEntry.IndexValue).EnumerateSequence
                            EnumerateSequence = EnumerateSequence & sEnumeratedSequence
                        End If
                    End If
                Case etIndex
                    Set oIndexEntry = FindIndex(oSubEntry.IndexName, oSubEntries)
                    If Not oIndexEntry Is Nothing Then
                        sEnumeratedSequence = oIndexEntry.EntryList(oIndexEntry.IndexValue).EnumerateSequence
                        EnumerateSequence = EnumerateSequence & sEnumeratedSequence
                    End If
                Case etText
                    sEnumeratedSequence = oSubEntry.EntryText
                    EnumerateSequence = EnumerateSequence & sEnumeratedSequence
            End Select
        Next
        iIndexPtr = iMaxIndexPtr
        Do
            bExit = False
            If oSubEntries(iIndexPtr).EntryType = etList Then
                Set oIndexEntry = FindIndex(oSubEntries(iIndexPtr).IndexName, oSubEntries)
                If oSubEntries(iIndexPtr) Is oIndexEntry Or oIndexEntry Is Nothing Then
                    oSubEntries(iIndexPtr).IndexValue = oSubEntries(iIndexPtr).IndexValue + 1
                    If oSubEntries(iIndexPtr).IndexValue = (oSubEntries(iIndexPtr).EntryList.Count + 1) Then
                        oSubEntries(iIndexPtr).IndexValue = 1
                        iIndexPtr = iIndexPtr - 1
                    Else
                        bExit = True
                    End If
                Else
                    iIndexPtr = iIndexPtr - 1
                End If
            Else
                iIndexPtr = iIndexPtr - 1
            End If
        Loop Until bExit Or iIndexPtr < 1
    Loop Until iIndexPtr < 1
    For Each oSubEntry In oSubEntries
        oSubEntry.IndexValue = 1
    Next
End Function

Private Function FindIndex(sIndex As String, oSubEntries As Collection) As Entry
    Dim oSubEntry As Entry
    
    If sIndex = "" Then
        Exit Function
    End If
    For Each oSubEntry In oSubEntries
        If oSubEntry.EntryType = etList Then
            If oSubEntry.IndexName = sIndex Then
                Set FindIndex = oSubEntry
                Exit For
            End If
        End If
    Next
End Function

Public Function EnumerateList() As Collection
    Dim iIndexPtr As Long
    Dim bExit As Boolean
    Dim iMaxIndexPtr As Long
    Dim oSubEntry As Entry
    Dim oIndexEntry As Entry
    Dim sEnumerateText As String
    Dim oSubEntries As Collection
    Dim oEnumeratedSequence As Collection
    
    On Error Resume Next
    
    Set oSubEntries = Me.EntryList
    
    If oSubEntries.Count = 0 Then
        Exit Function
    End If
    
    Set EnumerateList = New Collection
    
    For Each oSubEntry In oSubEntries
        Select Case oSubEntry.EntryType
            Case etSequence
                Set oEnumeratedSequence = oSubEntry.EnumerateSequence2
                AddCollection EnumerateList, oEnumeratedSequence
            Case etList
                Set oEnumeratedSequence = oSubEntry.EnumerateList
                AddCollection EnumerateList, oEnumeratedSequence
            Case etIndex
'                Set oIndexEntry = FindIndex(oSubEntry.IndexName, oSubEntries)
'                If Not oIndexEntry Is Nothing Then
'                    sEnumeratedSequence = oIndexEntry.EntryList(oIndexEntry.IndexValue).EnumerateSequence
'                    EnumerateSequence = EnumerateSequence & sEnumeratedSequence
'                End If
            Case etText
                Set oEnumeratedSequence = New Collection
                oEnumeratedSequence.Add oSubEntry.EntryText
                AddCollection EnumerateList, oEnumeratedSequence
        End Select
    Next
    
    Set Me.SetList = EnumerateList
End Function

Public Function EnumerateSequence2() As Collection
    Dim iIndexPtr As Long
    Dim bExit As Boolean
    Dim iMaxIndexPtr As Long
    Dim oSubEntry As Entry
    Dim oIndexEntry As Entry
    Dim sEnumerateText As String
    Dim oSubEntries As Collection
    Dim oEnumeratedSequence As Collection
    
    On Error Resume Next
    
    Set oSubEntries = Me.EntryList
    
    If oSubEntries.Count = 0 Then
        Exit Function
    End If

    iMaxIndexPtr = oSubEntries.Count
    Do
        For Each oSubEntry In oSubEntries
            Select Case oSubEntry.EntryType
                Case etSequence
                    Set oEnumeratedSequence = oSubEntry.EnumerateSequence2
                    AddCollection EnumerateList, oEnumeratedSequence
                Case etList
'                    If Not oSubEntry.Invisible Then
'                        Set oIndexEntry = FindIndex(oSubEntry.IndexName, oSubEntries)
'                        If oSubEntry Is oIndexEntry Or oIndexEntry Is Nothing Then
'                            sEnumeratedSequence = oSubEntry.EntryList(oSubEntry.IndexValue).EnumerateSequence
'                            EnumerateSequence = EnumerateSequence & sEnumeratedSequence
'                        Else
'                            sEnumeratedSequence = oSubEntry.EntryList(oIndexEntry.IndexValue).EnumerateSequence
'                            EnumerateSequence = EnumerateSequence & sEnumeratedSequence
'                        End If
'                    End If
                Case etIndex
'                    Set oIndexEntry = FindIndex(oSubEntry.IndexName, oSubEntries)
'                    If Not oIndexEntry Is Nothing Then
'                        sEnumeratedSequence = oIndexEntry.EntryList(oIndexEntry.IndexValue).EnumerateSequence
'                        EnumerateSequence = EnumerateSequence & sEnumeratedSequence
'                    End If
                Case etText
                    Set oEnumeratedSequence = New Collection
                    oEnumeratedSequence.Add oSubEntry.EntryText
                    AddCollection EnumerateList, oEnumeratedSequence
            End Select
        Next
        iIndexPtr = iMaxIndexPtr
        Do
            bExit = False
            If oSubEntries(iIndexPtr).EntryType = etList Then
                Set oIndexEntry = FindIndex(oSubEntries(iIndexPtr).IndexName, oSubEntries)
                If oSubEntries(iIndexPtr) Is oIndexEntry Or oIndexEntry Is Nothing Then
                    oSubEntries(iIndexPtr).IndexValue = oSubEntries(iIndexPtr).IndexValue + 1
                    If oSubEntries(iIndexPtr).IndexValue = (oSubEntries(iIndexPtr).EntryList.Count + 1) Then
                        oSubEntries(iIndexPtr).IndexValue = 1
                        iIndexPtr = iIndexPtr - 1
                    Else
                        bExit = True
                    End If
                Else
                    iIndexPtr = iIndexPtr - 1
                End If
            Else
                iIndexPtr = iIndexPtr - 1
            End If
        Loop Until bExit Or iIndexPtr < 1
    Loop Until iIndexPtr < 1
    For Each oSubEntry In oSubEntries
        oSubEntry.IndexValue = 1
    Next
    
    Set Me.SetList = EnumerateSequence2
End Function

Private Sub AddCollection(oCol1 As Collection, oCol2 As Collection)
    Dim vMember As Variant
    
    For Each vMember In oCol2
        oCol1.Add vMember
    Next
End Sub

Public Function FetchSetText() As String
    Dim vMember As Variant
    
    For Each vMember In Me.SetList
        FetchSetText = FetchSetText & vMember
    Next
End Function
