VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moPopulation(8) As New clsPicture

Private mlCanvasDC As Long

Public Property Let CanvasDC(ByVal lCanvasDC As Long)
    Dim lIndex As Long
    mlCanvasDC = lCanvasDC
    
    For lIndex = 0 To 8
        moPopulation(lIndex).CanvasDC = mlCanvasDC
    Next
End Property

Public Sub Render()
    Dim lIndex As Long
    
    For lIndex = 0 To 8
        moPopulation(lIndex).Render lIndex
    Next
End Sub

Public Sub MutateAll()
    Dim lIndex As Long
    
    For lIndex = 0 To 8
        moPopulation(lIndex).Mutate
    Next
End Sub

Public Sub SelectAndMutate(ByVal lIndex As Long)
    Dim oMember As clsPicture
    Dim lMemberIndex As Long
    
    Set oMember = moPopulation(lIndex)
    
    For lMemberIndex = 0 To 8
        If lMemberIndex <> lIndex Then
            Set moPopulation(lMemberIndex) = oMember.Clone
            moPopulation(lMemberIndex).Mutate
        End If
    Next
End Sub
