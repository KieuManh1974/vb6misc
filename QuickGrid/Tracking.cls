VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mlStamp As Long
Public moWatching As New Collection

Public Property Get CheckChanged() As Object
    Dim oObject As Object
    Dim lHighestStamp As Long
    
    lHighestStamp = mlStamp
    For Each oObject In moWatching
        If moWatching.moTracking.mlStamp > lHighestStamp Then
            lHighestStamp = moWatching.moTracking.mlStamp
            Set CheckChanged = oObject
        End If
    Next
    mlStamp = lHighestStamp
End Property

Public Sub Changed()
    mlStamp = mlStamp + 1
End Sub
