VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ListSetTypes
    Additive
    Multiplicative
    Textual
    ListReference
End Enum

Public Members As Collection
Public EnumeratedSet As Collection
Public SetEnumerated As Boolean
Public ListSetType As ListSetTypes
Public TextString As String
Public ListName As String
Public ListIndex As String
Public Invisible As Boolean
Public Reversed As Boolean

Public Function Text() As String
    Dim oEnumeratedMember As EnumeratedMember
    
    If Not SetEnumerated Then
        EnumerateSet
    End If
    For Each oEnumeratedMember In EnumeratedSet
        If Not oEnumeratedMember.Invisible Then
            Text = Text & oEnumeratedMember.TextMember
        End If
    Next
End Function

Public Sub EnumerateSet(Optional bReverse As Boolean)
    Dim oMember As ListSet
    Dim oEnumeratedMember As New EnumeratedMember
    Dim iIndex As Long

    Select Case ListSetType
        Case Additive
            EnumerateAdditiveSet bReverse
        Case Multiplicative
            EnumerateMultiplicativeSet bReverse
        Case Textual
            EnumerateTextualSet
        Case ListReference
            If IsNumeric(TextString) Then
                If Not bReverse Then
                    For iIndex = 1 To Val(TextString)
                        Set oEnumeratedMember = New EnumeratedMember
                        oEnumeratedMember.TextMember = CStr(iIndex)
                        EnumeratedSet.Add oEnumeratedMember
                    Next
                    SetEnumerated = True
                Else
                    For iIndex = Val(TextString) To 1 Step -1
                        Set oEnumeratedMember = New EnumeratedMember
                        oEnumeratedMember.TextMember = CStr(iIndex)
                        EnumeratedSet.Add oEnumeratedMember
                    Next
                    SetEnumerated = True
                End If
            Else
                For Each oMember In UniversalList
                    If oMember.ListName = TextString Then
                        If oMember.SetEnumerated Then
                            Set EnumeratedSet = oMember.EnumeratedSet
                        End If
                        SetEnumerated = True
                        Exit For
                    End If
                Next
            End If
    End Select
End Sub

Private Sub EnumerateAdditiveSet(Optional bReverse As Boolean)
    Dim oMember As ListSet
    Dim oSubEnumeratedMember As EnumeratedMember
    Dim oEnumeratedMember As EnumeratedMember
    Dim iMemberIndex As Integer
    Dim iStart As Long
    Dim iEnd As Long
    Dim iStep As Long
    
    If Not Reversed And Not bReverse Then
        iStart = 1
        iEnd = Members.Count
        iStep = 1
    Else
        iStart = Members.Count
        iEnd = 1
        iStep = -1
    End If
    
    For iMemberIndex = iStart To iEnd Step iStep
        Set oMember = Members(iMemberIndex)
        oMember.EnumerateSet bReverse Or Reversed
        For Each oSubEnumeratedMember In oMember.EnumeratedSet
            Set oEnumeratedMember = New EnumeratedMember
            oEnumeratedMember.TextMember = oSubEnumeratedMember.TextMember
            oEnumeratedMember.Invisible = oSubEnumeratedMember.Invisible
            EnumeratedSet.Add oEnumeratedMember
        Next
    Next
    SetEnumerated = True
End Sub

Private Sub EnumerateMultiplicativeSet(Optional bReverse As Boolean)
    Dim oMember As ListSet
    Dim iMemberIndex As Long
    Dim iMaxMember As Long
    Dim iCounter() As Long
    Dim iRedirectCounter() As Long
    Dim sThisItem As String
    Dim bFinished As Boolean
    Dim bExit As Boolean
    Dim iSubMemberIndex As Long
    Dim oEnumeratedMember As EnumeratedMember
    
    For Each oMember In Members
        oMember.EnumerateSet 'bReverse
    Next
    
    iMaxMember = Members.Count
    If iMaxMember > 0 Then
        ReDim iCounter(1 To iMaxMember) As Long
        ReDim iRedirectCounter(1 To iMaxMember) As Long
        For iMemberIndex = 1 To iMaxMember
            iCounter(iMemberIndex) = 1
            iRedirectCounter(iMemberIndex) = iMemberIndex
            If Members(iMemberIndex).ListName <> "" Then
                For iSubMemberIndex = 1 To iMemberIndex - 1
                    If Members(iMemberIndex).ListName = Members(iSubMemberIndex).ListName Then
                        iRedirectCounter(iMemberIndex) = iSubMemberIndex
                        Exit For
                    End If
                Next
            End If
        Next
        
        bFinished = False
        Do
            sThisItem = ""
            For iMemberIndex = 1 To iMaxMember
                If Members(iMemberIndex).EnumeratedSet.Count > 0 Then
                    If iCounter(iRedirectCounter(iMemberIndex)) <= Members(iMemberIndex).EnumeratedSet.Count Then
                        If Not Members(iMemberIndex).Invisible Then
                            sThisItem = sThisItem & Members(iMemberIndex).EnumeratedSet(iCounter(iRedirectCounter(iMemberIndex))).TextMember
                        End If
                    End If
                Else
                    If iRedirectCounter(iMemberIndex) <> iMemberIndex Then
                        If Not Members(iRedirectCounter(iMemberIndex)).EnumeratedSet(iCounter(iRedirectCounter(iMemberIndex))).Invisible Then
                            sThisItem = sThisItem & Members(iRedirectCounter(iMemberIndex)).EnumeratedSet(iCounter(iRedirectCounter(iMemberIndex))).TextMember
                        End If
                    End If
                End If
            Next
            Set oEnumeratedMember = New EnumeratedMember
            oEnumeratedMember.TextMember = sThisItem
            oEnumeratedMember.Invisible = False
            If Not bReverse Or EnumeratedSet.Count = 0 Then
                EnumeratedSet.Add oEnumeratedMember
            Else
                EnumeratedSet.Add oEnumeratedMember, , 1
            End If
            
            If Not Reversed Then
                iMemberIndex = 1
                Do
                    If iRedirectCounter(iMemberIndex) = iMemberIndex Then
                        iCounter(iMemberIndex) = iCounter(iMemberIndex) + 1
                        If iCounter(iMemberIndex) > Members(iMemberIndex).EnumeratedSet.Count Then
                            iCounter(iMemberIndex) = 1
                            iMemberIndex = iMemberIndex + 1
                        Else
                            iMemberIndex = 1
                        End If
                    Else
                        iMemberIndex = iMemberIndex + 1
                    End If
                Loop While iMemberIndex > 1 And iMemberIndex <= iMaxMember
                If iMemberIndex > iMaxMember Then
                    bFinished = True
                End If
            Else
                iMemberIndex = iMaxMember
                Do
                    If iRedirectCounter(iMemberIndex) = iMemberIndex Then
                        iCounter(iMemberIndex) = iCounter(iMemberIndex) + 1
                        If iCounter(iMemberIndex) > Members(iMemberIndex).EnumeratedSet.Count Then
                            iCounter(iMemberIndex) = 1
                            iMemberIndex = iMemberIndex - 1
                        Else
                            iMemberIndex = iMaxMember
                        End If
                    Else
                        iMemberIndex = iMemberIndex - 1
                    End If
                Loop While iMemberIndex < iMaxMember And iMemberIndex >= 1
                If iMemberIndex < 1 Then
                    bFinished = True
                End If
            End If
        Loop While Not bFinished
    End If
    SetEnumerated = True
End Sub

Private Sub EnumerateTextualSet()
    Dim oEnumeratedMember As New EnumeratedMember
    oEnumeratedMember.TextMember = TextString
    EnumeratedSet.Add oEnumeratedMember
    SetEnumerated = True
End Sub

Private Sub Class_Initialize()
    Set Members = New Collection
    Set EnumeratedSet = New Collection
End Sub
