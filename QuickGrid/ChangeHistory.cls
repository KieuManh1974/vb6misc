VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChangeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msChanges As Variant

Private Sub Class_Initialize()
    msChanges = Array()
End Sub

Public Sub Log(sText As String, ByVal lStart As Long, ByVal lLength As Long)
    ReDim Preserve msChanges(UBound(msChanges) + 1)
    msChanges(UBound(msChanges)) = Array(sText, lStart, lLength)
End Sub

Public Function Undo() As Variant
    If UBound(msChanges) >= 0 Then
        Undo = msChanges(UBound(msChanges))
        If UBound(msChanges) = 0 Then
            msChanges = Array()
        Else
            ReDim Preserve msChanges(UBound(msChanges) - 1)
        End If
    End If
End Function

Public Sub Clear()
    msChanges = Array()
End Sub
