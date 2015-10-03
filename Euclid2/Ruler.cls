VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ruler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IShape

Private mssShapeState As ShapeStates
Private moCanvasSurface As Canvas

Public LineStyle As Long ' LineStyles
Public LineLength As Double

Public Positions As New Collection

Private mdAngle As Double

Private Sub Class_Initialize()
    Dim oConnector As New Connector
    mdAngle = Rnd * PI2
    
    Positions.Add oConnector
    
    LineStyle = lsLeftDotted + lsRightSolid
End Sub

Private Function IShape_Copy() As IShape

End Function

Private Sub IShape_DragMe(oVector As Vector)

End Sub

Private Sub IShape_Move(oVector As Vector, ByVal Shift As Integer)
    Dim oNewLine As Vector
    
    If Positions.Count = 1 Then
        If Shift = 0 Then
            Set oNewLine = oVector.Subtract(Positions(1).Position)
            
            mdAngle = Atn2(oNewLine)
            
        End If
    End If
End Sub

Private Sub IShape_ReleaseMe()

End Sub

Private Sub IShape_Render()
    Dim lColour As Long
    Dim oLineVectorNormal As Vector
    Dim nHalfScreenWidth As Single
    Dim oLinePolar As Vector
    
    Select Case mssShapeState
        Case ShapeStates.ssNormal
            lColour = euColourNormal
        Case ShapeStates.ssSelected, ShapeStates.ssDragged
            lColour = euColourSelected
        Case ShapeStates.ssRemove
            lColour = moCanvasSurface.Surface.BackColor
    End Select
    
    'Set oLineVectorNormal = Position2.Subtract(Positions(1).Position).Normal
    
    nHalfScreenWidth = moCanvasSurface.Surface.ScaleWidth / 2
    
    If Positions.Count = 1 Then
        Set oLinePolar = New Vector
        oLinePolar.CreatePolar nHalfScreenWidth, mdAngle
        
        ' Left
        If (LineStyle Mod 3) = lsLeftDotted Then
            moCanvasSurface.Surface.DrawStyle = vbDot
            moCanvasSurface.Surface.Line (Positions(1).Position.X, Positions(1).Position.Y)-Step(-oLinePolar.X, -oLinePolar.Y), lColour
        ElseIf (LineStyle Mod 3) = lsLeftSolid Then
            moCanvasSurface.Surface.DrawStyle = vbSolid
            moCanvasSurface.Surface.Line (Positions(1).Position.X, Positions(1).Position.Y)-Step(-oLinePolar.X, -oLinePolar.Y), lColour
        End If
        
        ' Right
        If 9 * ((LineStyle \ 9) Mod 3) = lsRightDotted Then
            moCanvasSurface.Surface.DrawStyle = vbDot
            moCanvasSurface.Surface.Line (Positions(1).Position.X, Positions(1).Position.Y)-Step(oLinePolar.X, oLinePolar.Y), lColour
        ElseIf 9 * ((LineStyle \ 9) Mod 3) = lsRightSolid Then
            moCanvasSurface.Surface.DrawStyle = vbSolid
            moCanvasSurface.Surface.Line (Positions(1).Position.X, Positions(1).Position.Y)-Step(oLinePolar.X, oLinePolar.Y), lColour
        End If
    Else
'        ' Left
'        If (LineStyle Mod 3) = lsLeftDotted Then
'            moCanvasSurface.Surface.DrawStyle = vbDot
'            moCanvasSurface.Surface.Line (Position.X, Position.Y)-(oLineVectorNormal.Scaling(-nHalfScreenWidth).X, oLineVectorNormal.Scaling(-nHalfScreenWidth).Y), lColour
'        ElseIf (LineStyle Mod 3) = lsLeftSolid <> 0 Then
'            moCanvasSurface.Surface.DrawStyle = vbSolid
'            moCanvasSurface.Surface.Line (Position.X, Position.Y)-(oLineVectorNormal.Scaling(-nHalfScreenWidth).X, oLineVectorNormal.Scaling(-nHalfScreenWidth).Y), lColour
'        End If
'
'        ' Middle
'        If 3 * ((LineStyle \ 3) Mod 3) = lsMidDotted Then
'            moCanvasSurface.Surface.DrawStyle = vbDot
'            moCanvasSurface.Surface.Line (Position.X, Position.Y)-(Position2.X, Position2.Y), lColour
'        ElseIf 3 * ((LineStyle \ 3) Mod 3) = lsMidSolid Then
'            moCanvasSurface.Surface.DrawStyle = vbSolid
'            moCanvasSurface.Surface.Line (Position.X, Position.Y)-(Position2.X, Position2.Y), lColour
'        End If
'
'        ' Right
'        If 9 * ((LineStyle \ 9) Mod 3) = lsRightDotted Then
'            moCanvasSurface.Surface.DrawStyle = vbDot
'            moCanvasSurface.Surface.Line (Position2.X, Position2.Y)-(oLineVectorNormal.Scaling(nHalfScreenWidth).X, oLineVectorNormal.Scaling(nHalfScreenWidth).Y), lColour
'        ElseIf 9 * ((LineStyle \ 9) Mod 3) = lsRightSolid Then
'            moCanvasSurface.Surface.DrawStyle = vbSolid
'            moCanvasSurface.Surface.Line (Position2.X, Position2.Y)-(oLineVectorNormal.Scaling(nHalfScreenWidth).X, oLineVectorNormal.Scaling(nHalfScreenWidth).Y), lColour
'        End If
    End If
End Sub

Private Sub IShape_SelectMe(oVector As Vector)
    Dim dDist As Double
    Dim oLine As Vector
    Dim oLine2 As Vector
    Dim oDiff As Vector
    Dim oDiff2 As Vector
    Dim dSide As Double
    Dim dSide2 As Double
    Dim oPolar As Vector
    
    If Positions.Count = 1 Then
        Set oPolar = New Vector
        oPolar.CreatePolar 1, mdAngle
        
        Set oDiff = oVector.Subtract(Positions(1).Position)
        
        dDist = Abs(oDiff.Cross(oPolar))
        
        If dDist <= TouchRadius Then
            If mssShapeState <> ssSelected Then
                IShape_State = ssSelected
            End If
        Else
            If mssShapeState = ssSelected Then
                IShape_State = ssNormal
            End If
        End If
        
    Else
'        Set oLine = Position2.Subtract(Position)
'        Set oLine2 = Position.Subtract(Position2)
'
'        Set oDiff = oVector.Subtract(Position)
'        Set oDiff2 = oVector.Subtract(Position2)
'
'        dDist = Abs(oDiff.Cross(oLine) / oLine.Size)
'        dSide = oDiff.Dot(oLine)
'        dSide2 = oDiff2.Dot(oLine)
'
'        If dSide >= 0 And dSide2 <= 0 Then
'            If dDist <= TouchRadius Then
'                If mssShapeState <> ssSelected Then
'                    IShape_State = ssSelected
'                End If
'            Else
'                If mssShapeState = ssSelected Then
'                    IShape_State = ssNormal
'                End If
'            End If
'        End If
    End If
End Sub

Private Property Set IShape_CanvasSurface(oCanvasSurface As Canvas)
    Set moCanvasSurface = oCanvasSurface
End Property

Private Property Get IShape_CanvasSurface() As Canvas
    Set IShape_CanvasSurface = moCanvasSurface
End Property

Private Property Let IShape_State(ByVal ssShapeState As ShapeStates)
    mssShapeState = ssShapeState
    IShape_Render
End Property

Private Property Get IShape_State() As ShapeStates
    IShape_State = mssShapeState
End Property

Private Sub IShape_Translate(oVector As Vector)

End Sub
