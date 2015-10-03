Attribute VB_Name = "Module1"
Option Explicit

Public Sub main()
    Dim x As Long
    
    For x = 0 To 9
        Debug.Print MapRange(x, 1, 9, 13)
    Next
End Sub

Public Function MapRange(ByVal lValue As Long, ByVal lSourceOffset As Long, ByVal lTargetStart As Long, ByVal lTargetEnd As Long) As Long
    Dim lRange As Long
    
    lRange = lTargetEnd - lTargetStart + 1
    MapRange = ((((lValue - lSourceOffset) Mod lRange) + lRange) Mod lRange) + lTargetStart
End Function

