VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LeftSide As Double
Public RightSide As Double
Public TopSide As Double
Public BottomSide As Double
Public Canvas As Form

Private Sides() As Variant
Private CapturedX() As Integer

Public Property Get Captured(ByVal iIndex As Integer) As Integer
    Captured = CapturedX(iIndex)
End Property

Public Property Let Captured(ByVal iIndex As Integer, ByVal iValue As Integer)
    CapturedX(iIndex) = iValue
End Property

Public Property Get Side(iSideIndex As Integer) As Variant
    Side = Sides(iSideIndex)
End Property

Public Property Let Side(iSideIndex As Integer, vValue As Variant)
    Sides(iSideIndex) = vValue
End Property

Public Property Let SideValue(iSideIndex As Integer, iValueIndex As Integer, sValue As Single)
    Sides(iSideIndex)(iValueIndex) = sValue
End Property

Public Property Get SideCount() As Integer
    SideCount = UBound(Sides)
End Property

Public Sub DrawWall(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    Dim subound As Integer
    subound = UBound(Sides) + 4
    ReDim Preserve Sides(subound) As Variant
    Sides(subound - 3) = Array(x1, y1, x2, y1)
    Sides(subound - 2) = Array(x2, y1, x2, y2)
    Sides(subound - 1) = Array(x2, y2, x1, y2)
    Sides(subound) = Array(x1, y2, x1, y1)
    ReDim CapturedX(UBound(CapturedX) + 4) As Integer
    
    Canvas.FillStyle = vbDownwardDiagonal
    Canvas.Line (x1, y1)-(x2, y2), vbRed, B
    Canvas.FillStyle = vbFSTransparent
End Sub

Public Sub DrawPartition(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    ReDim Preserve Sides(UBound(Sides) + 1) As Variant
    Sides(UBound(Sides)) = Array(x1, y1, x2, y2)
    ReDim CapturedX(UBound(CapturedX) + 1) As Integer
    
    Canvas.Line (x1, y1)-(x2, y2)
    Canvas.Circle (x1, y1), 2
    Canvas.Circle (x2, y2), 2
End Sub

Public Sub Polygon(cx As Double, cy As Double, w As Double, s As Integer)
    Dim stepper As Integer
    Dim angle As Double
    Dim oangle As Double
    Dim pi As Double
    Dim subound As Integer
    
    subound = UBound(Sides) + s
    ReDim Preserve Sides(subound) As Variant
    
    pi = 4 * Atn(1)
    For stepper = 0 To s - 1
        oangle = (stepper - 1) * 2 * pi / s
        angle = stepper * 2 * pi / s
        Canvas.Line (cx + w * Cos(angle), cy + w * Sin(angle))-(cx + w * Cos(oangle), cy + w * Sin(oangle)), vbCyan
        Sides(subound - stepper) = Array(cx + w * Cos(angle), cy + w * Sin(angle), cx + w * Cos(oangle), cy + w * Sin(oangle))
    Next
    ReDim CapturedX(UBound(CapturedX) + s) As Integer
End Sub

Private Sub Class_Initialize()
    ReDim CapturedX(0) As Integer
    ReDim Sides(0) As Variant
End Sub
