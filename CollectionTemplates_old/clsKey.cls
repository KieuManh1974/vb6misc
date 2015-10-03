VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsIKey

Private msPartKey As String
Private moSubPartKeys() As clsKey
Private mlSubPartKeyCount As Long
Private mbAssigned As Boolean

Public Item As Variant
Public Parent As clsKey

Private Enum Operations
    opSub
    opSuper
    opEqual
    opOverlap
End Enum

Private mopOperation As Operations
Private msNewPartKey As String

Public Property Get Top() As clsKey
    If Not Parent Is Nothing Then
        Set Top = Parent.Top
    Else
        Set Top = Me
    End If
End Property

Public Property Get PartKey() As String
    PartKey = msPartKey
End Property

Public Property Let PartKey(ByVal sPartKey As String)
    msPartKey = sPartKey
End Property

Public Property Get NewPartKey() As String
    NewPartKey = msNewPartKey
End Property

Public Property Get Operation() As Long
    Operation = mopOperation
End Property

Public Property Let Assigned(ByVal bAssigned As Boolean)
    mbAssigned = bAssigned
End Property

Public Property Get Assigned() As Boolean
    Assigned = mbAssigned
End Property

Public Function AddKey(ByVal sKey As String) As clsKey
    Dim oKey As clsKey
    Dim sRightKey As String
    Dim lOverlap As Long
    Dim oNewKey As clsKey
    Dim oTop As clsKey
    
    Set oTop = Top
    Set oKey = oTop.FindPartialKey(sKey)
    If Not oKey Is Nothing Then
        Select Case oKey.Operation
            Case Operations.opEqual
                oKey.Assigned = True
                Set AddKey = oKey
            Case Operations.opSuper
                Set AddKey = New clsKey
                AddKey.PartKey = Mid$(oKey.NewPartKey, Len(oKey.PartKey) + 1)
                AddKey.Assigned = True
                Set AddKey.Parent = oKey
                oKey.AddChild AddKey
            Case Operations.opSub
                oKey.PartKey = Mid$(oKey.PartKey, Len(oKey.NewPartKey) + 1)
                Set AddKey = New clsKey
                AddKey.PartKey = oKey.NewPartKey
                AddKey.AddChild oKey
                Set AddKey.Parent = oKey.Parent
                AddKey.Assigned = True
                Set oKey.Parent = AddKey
            Case Operations.opOverlap
                lOverlap = StringOverlapPos(oKey.PartKey, oKey.NewPartKey)
                oKey.PartKey = Mid$(oKey.PartKey, lOverlap + 1)
                Set AddKey = New clsKey
                AddKey.PartKey = Left$(oKey.NewPartKey, lOverlap)
                AddKey.AddChild oKey
                Set AddKey.Parent = oKey.Parent
                Set oKey.Parent = AddKey
                Set oNewKey = AddKey.NewSubKey
                oNewKey.Assigned = True
                oNewKey.PartKey = Mid$(oKey.NewPartKey, lOverlap + 1)
        End Select
    Else
        If mlSubPartKeyCount = 0 Then
            If Not mbAssigned Then
                msPartKey = sKey
                mbAssigned = True
                Set AddKey = Me
            Else
                oTop.PartKey = Mid$(oTop.PartKey, Len(oTop.NewPartKey) + 1)
                Set AddKey = New clsKey
                AddKey.PartKey = oTop.NewPartKey
                AddKey.AddChild oTop
                Set AddKey.Parent = oTop.Parent
                AddKey.Assigned = True
                Set oTop.Parent = AddKey
                Set oNewKey = AddKey.NewSubKey
                oNewKey.Assigned = True
                oNewKey.PartKey = sKey
            End If
        End If
    End If
End Function

Public Property Get Key(ByVal sKey) As clsKey
    Dim oKey As clsKey
    
    Set oKey = Top.FindPartialKey(sKey)
    Select Case oKey.Operation
        Case Operations.opEqual
            If oKey.Assigned Then
                Set Key = oKey
            End If
    End Select
End Property

Public Function FindPartialKey(ByVal sKey As String) As clsKey
    Dim lOverlap As Long
    Dim lLengthKey As Long
    Dim lLengthMyKey As Long
    Dim lIndex As Long
    Dim oSubPartKey As clsKey
    Dim sRightKey As String
    
    Set FindPartialKey = Me
    
    lLengthMyKey = Len(msPartKey)
    lLengthKey = Len(sKey)
        
    lOverlap = StringOverlapPos(msPartKey, sKey)
    If Len(sKey) > Len(msPartKey) Then
        If lOverlap = 0 Then
            Set FindPartialKey = Nothing
            Exit Function
        ElseIf lOverlap = lLengthMyKey Then
            ' Check children
            sRightKey = Mid$(sKey, lLengthMyKey + 1)
            For lIndex = 0 To mlSubPartKeyCount - 1
                Set oSubPartKey = moSubPartKeys(lIndex).FindPartialKey(sRightKey)
                If Not oSubPartKey Is Nothing Then
                    Set FindPartialKey = oSubPartKey
                    Exit For
                End If
            Next
            mopOperation = opSuper
            msNewPartKey = sKey
        Else
        End If
    ElseIf Len(sKey) < Len(msPartKey) Then
        If lOverlap = 0 Then
            Set FindPartialKey = Nothing
            Exit Function
        ElseIf lOverlap = lLengthKey Then
            mopOperation = opSub
            msNewPartKey = sKey
        Else
            Exit Function
        End If
    Else
        If lOverlap = 0 Then
            Set FindPartialKey = Nothing
            Exit Function
        Else
            If lOverlap = lLengthMyKey Then
                mopOperation = opEqual
                msNewPartKey = sKey
            Else
                mopOperation = opOverlap
                msNewPartKey = sKey
            End If
        End If
    End If
End Function

Private Function StringOverlapPos(ByVal sString1 As String, ByVal sString2 As String) As Long
    Dim lIndex As Long
    Dim lLength1 As Long
    Dim lLength2 As Long
    
    lLength1 = Len(sString1)
    lLength2 = Len(sString2)
    
    lIndex = 1
    Do
        If lIndex > lLength1 Then
            StringOverlapPos = lLength1
            Exit Do
        End If
        If lIndex > lLength2 Then
            StringOverlapPos = lLength2
            Exit Do
        End If
        
        If Mid$(sString1, lIndex, 1) <> Mid$(sString2, lIndex, 1) Then
            StringOverlapPos = lIndex - 1
            Exit Do
        End If
        lIndex = lIndex + 1
    Loop
End Function

Public Function NewSubKey() As clsKey
    Set NewSubKey = New clsKey
    
    ReDim Preserve moSubPartKeys(mlSubPartKeyCount)
    Set moSubPartKeys(mlSubPartKeyCount) = NewSubKey
    mlSubPartKeyCount = mlSubPartKeyCount + 1
End Function

Public Function AddChild(ByVal oChild As clsKey)
    ReDim Preserve moSubPartKeys(mlSubPartKeyCount)
    Set moSubPartKeys(mlSubPartKeyCount) = oChild
    mlSubPartKeyCount = mlSubPartKeyCount + 1
End Function


Private Function clsIKey_AddItem(ByVal sKey As String, ByVal vItem As Variant)
    Dim oKey As clsKey
    
    Set oKey = AddKey(sKey)
    oKey.Item = vItem
End Function

Private Property Get clsIKey_Item(ByVal sKey As Variant) As Variant
    Dim oKey As clsKey
    
    Set oKey = Key(sKey)
    oKey.Item = oKey.Item
End Property

Private Function clsIKey_RemoveItem(ByVal sKey As String) As clsKey

End Function
