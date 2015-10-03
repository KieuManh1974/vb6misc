VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Point
    X As Long
    Y As Long
    Colour As Long
    Width As Long
End Type

Private Type RGB
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
End Type

Private mpPoints() As Point
Private mlCanvasDC As Long

Public Property Let CanvasDC(ByVal lCanvasDC As Long)
    mlCanvasDC = lCanvasDC
End Property

Public Function SetPoint(ByVal lIndex As Long, ByVal lX As Long, ByVal lY As Long, ByVal lColour As Long, ByVal lWidth As Long)
    If lIndex > UBound(mpPoints) Then
        ReDim Preserve mpPoints(lIndex)
    End If
    With mpPoints(lIndex)
        .X = lX
        .Y = lY
        .Colour = lColour
        .Width = lWidth
    End With
End Function

Public Sub Render(lPosition)
    Dim lIndex As Long
    Dim lColumn As Long
    Dim lRow As Long
    
    Dim lX As Long
    Dim lY As Long
    
    lColumn = glCellWidth * (lPosition Mod 3)
    lRow = glCellWidth * (lPosition \ 3)
    
    For lIndex = 0 To UBound(mpPoints)
        For lY = 0 To mpPoints(lIndex).Width - 1
            For lX = 0 To mpPoints(lIndex).Width - 1
                SetPixelV mlCanvasDC, mpPoints(lIndex).X + lColumn + 50 + lX, mpPoints(lIndex).Y + lRow + 50 + lY, mpPoints(lIndex).Colour
            Next
        Next
    Next
End Sub

Public Sub Mutate()
    Dim lTimes As Long

    For lTimes = 0 To Rnd * 1000
        Select Case Rnd
            Case Is < 0.001
                RemovePoint
            Case Is < 0.002
                AddPoint
            Case Is < 0.1
                MovePoint
            Case Else
                ChangeShade
        End Select
    Next
End Sub

Private Sub RemovePoint()
    Dim lPointIndex As Long
    Dim lIndex As Long
    
    If UBound(mpPoints) = 0 Then
        Exit Sub
    End If
    
    lPointIndex = Int(Rnd * (UBound(mpPoints) + 1))
    For lIndex = lPointIndex To UBound(mpPoints) - 1
        mpPoints(lIndex) = mpPoints(lIndex + 1)
    Next
    ReDim Preserve mpPoints(UBound(mpPoints) - 1)
End Sub

Private Sub AddPoint()
    Dim lShade As Long
    
    ReDim Preserve mpPoints(UBound(mpPoints) + 1)
    With mpPoints(UBound(mpPoints))
        .X = Rnd * glCellWidth * 0.75
        .Y = Rnd * glCellWidth * 0.75
        .Colour = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
        .Width = Int(Rnd * 20) + 1
    End With
End Sub

Private Sub MovePoint()
    Dim lRnd As Long
    Dim lPointIndex As Long
    
    lPointIndex = Int(Rnd * (UBound(mpPoints) + 1))
    With mpPoints(lPointIndex)
        .X = Random(.X, 0, glCellWidth * 0.75, 3)
        .Y = Random(.Y, 0, glCellWidth * 0.75, 3)
        .Width = Random(.Width, 1, 100, 3)
    End With
End Sub

Private Sub ChangeShade()
    Dim lRGB As RGB
    Dim lShade As Long
    Dim lPointIndex As Long
    Dim lRnd As Long
    
    lPointIndex = Int(Rnd * (UBound(mpPoints) + 1))
    With mpPoints(lPointIndex)
        CopyMem lRGB, ByVal (VarPtr(.Colour)), 4&
        lRGB.Red = Random(lRGB.Red, 0, 255, 3)
        lRGB.Green = Random(lRGB.Green, 0, 255, 3)
        lRGB.Blue = Random(lRGB.Blue, 0, 255, 3)
        .Colour = RGB(lRGB.Red, lRGB.Green, lRGB.Blue)
    End With
End Sub

Private Sub Class_Initialize()
    Dim lIndex As Long
    
    ReDim mpPoints(0)
    For lIndex = 0 To 300
        AddPoint
    Next
End Sub

Public Function Clone() As clsPicture
    Dim lIndex As Long
    
    Set Clone = New clsPicture
    For lIndex = 0 To UBound(mpPoints)
        With mpPoints(lIndex)
            Clone.SetPoint lIndex, .X, .Y, .Colour, .Width
        End With
    Next
    Clone.CanvasDC = mlCanvasDC
End Function

Private Function Random(ByVal lValue As Long, ByVal lFloor As Long, ByVal lCeiling As Long, ByVal lRange As Long)
    Dim lRnd As Long

    lRnd = Int(Rnd * lRange) - lRange \ 2
    While (lValue + lRnd) < lFloor Or (lValue + lRnd) > lCeiling
        lRnd = Int(Rnd * lRange) - lRange \ 2
    Wend
    lValue = lValue + lRnd
    Random = lValue
End Function
