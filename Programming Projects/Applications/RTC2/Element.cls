VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ReverseMultiplied As Boolean
Public Members As New Collection

Public Sub CreateElement(oTree As ParseTree)
    Dim oSubTree As ParseTree
    Dim oPart As Part
    Dim lLinkCheck As Long
    Dim oSet As SetDef
    
    If oTree(1).Index = 1 Then
        ReverseMultiplied = True
    End If
    
    For Each oSubTree In oTree(2).SubTree
        Set oPart = New Part
        Select Case oSubTree.Index
            Case 1 ' assignment
                oPart.PartType = ptAssignment
                oPart.VarName = oSubTree(1)(1).Text
                Set oPart.PartSetDef = New SetDef
                oPart.PartSetDef.CreateSet oSubTree(1)(2)
                AllSets.Add CStr(oPart.VarName), oPart.PartSetDef
            Case 2 ' set var
                oPart.PartType = ptSetVar
                oPart.VarName = oSubTree.Text
                If AllSets.Exists(CStr(oPart.VarName)) Then
                    Set oPart.PartSetDef = AllSets.Item(CStr(oPart.VarName))
                End If
            Case 3 ' reference
                oPart.PartType = ptParallelVar
                oPart.Parallel = True
                oPart.VarName = oSubTree(1)(1).Text
                If AllSets.Exists(CStr(oPart.VarName)) Then
                    Set oPart.PartSetDef = AllSets.Item(CStr(oPart.VarName))
                End If
                Set oSet = New SetDef
                oSet.CreateSet oSubTree(1)(2)
                Set oPart.ReferenceSet = oSet
                For lLinkCheck = 1 To Members.Count
                    If Members(lLinkCheck).Parallel = False And Members(lLinkCheck).VarName = oPart.VarName Then
                        oPart.MemberLink = lLinkCheck
                        Exit For
                    End If
                Next
            Case 4 ' parallel var
                oPart.PartType = ptParallelVar
                oPart.Parallel = True
                oPart.VarName = oSubTree.Text
                If AllSets.Exists(CStr(oPart.VarName)) Then
                    Set oPart.PartSetDef = AllSets.Item(CStr(oPart.VarName))
                End If
                For lLinkCheck = 1 To Members.Count
                    If Members(lLinkCheck).Parallel = False And Members(lLinkCheck).VarName = oPart.VarName Then
                        oPart.MemberLink = lLinkCheck
                        Exit For
                    End If
                Next
            Case 5 ' range
                oPart.LeadingZeroes = Left$(oSubTree(1)(1).Text, 1) = "0"
                oPart.RangeStart = Val(oSubTree(1)(1).Text)
                If oSubTree(1)(2).Index = 1 Then
                    oPart.PartType = ptRange
                    oPart.RangeEnd = Val(oSubTree(1)(2).Text)
                Else
                    oPart.PartType = ptNumber
                End If
            Case 6 ' set
                oPart.PartType = ptSetDef
                Set oPart.PartSetDef = New SetDef
                oPart.PartSetDef.CreateSet oSubTree(1)
            Case 7 ' escape
                oPart.PartType = ptEscape
                oPart.Character = oSubTree.Text
            Case 8 ' other
                oPart.PartType = ptChar
                Select Case oSubTree.Text
                    Case "^"
                        oPart.Character = vbCrLf
                    Case "%"
                        oPart.Character = vbTab
                    Case Else
                        oPart.Character = oSubTree.Text
                End Select
            Case 9 ' Blank
                oPart.PartType = ptBlank
        End Select
        Members.Add oPart
    Next
End Sub

Public Property Get Text(ByVal lIndex As Long) As String
    Dim lIndeces() As Long
    Dim lPosition As Long
    
    ReDim lIndeces(1 To Members.Count)

    If Not ReverseMultiplied Then
        For lPosition = 1 To Members.Count
            If Members(lPosition).MemberLink = 0 Then
                lIndeces(lPosition) = lIndex Mod Members(lPosition).SetSize
                lIndex = lIndex \ Members(lPosition).SetSize
                If Members(lPosition).ReferenceSet Is Nothing Then
                    Text = Text & Members(lPosition).Text(lIndeces(lPosition))
                Else
                    Text = Text & Members(lPosition).ReferenceSet.Text(lIndeces(lPosition))
                End If
            Else
                If Members(lPosition).ReferenceSet Is Nothing Then
                    Text = Text & Members(Members(lPosition).MemberLink).Text(lIndeces(Members(lPosition).MemberLink))
                Else
                    Text = Text & Members(lPosition).ReferenceSet.Text(lIndeces(Members(lPosition).MemberLink))
                End If
            End If
        Next
    Else
        For lPosition = Members.Count To 1 Step -1
            lIndeces(lPosition) = lIndex Mod Members(lPosition).SetSize
            lIndex = lIndex \ Members(lPosition).SetSize
        Next
        
        For lPosition = Members.Count To 1 Step -1
            If Members(lPosition).MemberLink = 0 Then
                Text = Members(lPosition).Text(lIndeces(lPosition)) & Text
            Else
                Text = Members(Members(lPosition).MemberLink).Text(lIndeces(Members(lPosition).MemberLink)) & Text
            End If
        Next
    End If
End Property

Public Property Get SetSize() As Long
    Dim oPart As Part
    
    SetSize = 1
    For Each oPart In Members
        If oPart.MemberLink = 0 Then
            SetSize = SetSize * oPart.SetSize
        End If
    Next
End Property
