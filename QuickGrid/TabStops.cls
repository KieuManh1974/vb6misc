VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabStops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moTabStops As Variant
Private mlDefaultWidth As Long
Public mlMinimumTabWidth As Long

Public Sub Clear()
    moTabStops = Array(0)
End Sub
 
Public Property Let TabStop(ByVal lStopIndex As Long, ByVal lPosition As Long)
    Dim lUbound As Long
    Dim lFillIndex As Long
    Dim lOffset As Long
    
    lUbound = UBound(moTabStops)
    If lStopIndex > lUbound Then
        ReDim Preserve moTabStops(lStopIndex)
    End If
    
    For lFillIndex = lUbound + 1 To lStopIndex - 1
        moTabStops(lFillIndex) = moTabStops(lFillIndex - 1) + mlDefaultWidth
    Next
    
    If lStopIndex < lUbound Then
        lOffset = lPosition - moTabStops(lStopIndex)
        For lFillIndex = lStopIndex To UBound(moTabStops)
            moTabStops(lFillIndex) = moTabStops(lFillIndex) + lOffset
        Next
    Else
        moTabStops(lStopIndex) = lPosition
    End If
End Property

Public Property Get TabStop(ByVal lStopIndex As Long) As Long
    If lStopIndex <= UBound(moTabStops) Then
        TabStop = moTabStops(lStopIndex)
    Else
        TabStop = lStopIndex * mlDefaultWidth
    End If
End Property

Public Function GetTabStopForPosition(ByVal lPosition As Long)
    Dim lStopIndex As Long
    
    If lPosition < 0 Then
        GetTabStopForPosition = -1
        Exit Function
    End If
    
    While TabStop(lStopIndex) < lPosition
        lStopIndex = lStopIndex + 1
    Wend
    GetTabStopForPosition = lStopIndex - 1
End Function

Public Property Let DefaultWidth(ByVal lWidth As Long)
    mlDefaultWidth = lWidth
    moTabStops = Array(0)
End Property

Public Property Get DefaultWidth() As Long
    DefaultWidth = mlDefaultWidth
End Property

Private Sub Class_Initialize()
    moTabStops = Array(0)
End Sub
