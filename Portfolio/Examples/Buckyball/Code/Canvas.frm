VERSION 5.00
Begin VB.Form Canvas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H007C7303&
   Caption         =   "Buckyball"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   600
      Top             =   840
   End
End
Attribute VB_Name = "Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Logical Brush (or Pattern)
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

' Brush Styles
Const BS_SOLID = 0
Const BS_NULL = 1
Const BS_HOLLOW = BS_NULL
Const BS_HATCHED = 2
Const BS_PATTERN = 3
Const BS_INDEXED = 4
Const BS_DIBPATTERN = 5
Const BS_DIBPATTERNPT = 6
Const BS_PATTERN8X8 = 7
Const BS_DIBPATTERN8X8 = 8


' PolyFill() Modes
Const ALTERNATE = 1
Const WINDING = 2
Const POLYFILL_LAST = 2

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Cube As Entity
Private Football As Entity

Private PitchRate As Double
Private RollRate As Double
Private PitchAbs As Double
Private RollAbs As Double
Private PositionAbs As Vector
Private Speed As Double
Private MyAxis As Vector
Private MyRightSide As Vector
Private MyAngle As Double

Private pi As Double
Private roll As Double
Private pitch As Double
Private mewidth As Single
Private meheight As Single


Private Sub InitialiseFootball()
    Dim hexface As VectorSet
    Dim pentface As VectorSet
    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim cl As Double
    Dim p As Integer
    Dim angle As Double
    Dim newpos As Vector
    
    ReDim hexface.lineref(0 To 5) As Integer
    ReDim pentface.lineref(0 To 4) As Integer
    
    With Football
        With .positional
            .axial.y = 1
            .position.z = 0
            .rightside.x = 1
            .down.z = 1
        End With
        
        ReDim .points(0 To 59) As Vector
        ReDim .points2d(0 To 59) As Vector2d
        ReDim .facets(0 To 31) As VectorSet
        
        cx = -Cos(120 * Atn(1) * 8 / 360)
        cy = Cos(120 * Atn(1) * 8 / 360) * (1 + Cos(72 * Atn(1) * 8 / 360)) / Sin(72 * Atn(1) * 8 / 360)
        cz = Sqr(1 - cx * cx - cy * cy)
        
        cl = Sqr(cx * cx + cy * cy)
        

        ReDim .lines(0 To 94) As LineSet
        Dim xaxis As Vector
        Dim yaxis As Vector
        Dim v1 As Vector
        Dim v2 As Vector
        Dim v3 As Vector
        Dim ps As Integer
        
        Dim iSpoke As Integer
        
        ' First pentagon
        For iSpoke = 0 To 4
            angle = iSpoke * Atn(1) * 8 / 5
            .points(PointOf(0, iSpoke)) = VectorOf(100 * Cos(angle), 100 * Sin(angle), 0) ' hexagon
            .points(PointOf(1, iSpoke)) = Add(.points(PointOf(0, iSpoke)), VectorOf(100 * cl * Cos(angle), 100 * cl * Sin(angle), 100 * cz)) ' pentagon mids
        Next
        
        ' Next hexagons
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(1, iSpoke)), .points(PointOf(0, iSpoke)))
            v2 = Subs(.points(PointOf(1, iSpoke + 1)), .points(PointOf(0, iSpoke + 1)))
            
            .points(PointOf(2, iSpoke, 1)) = Add(.points(PointOf(1, iSpoke)), v2)
            .points(PointOf(2, iSpoke + 1, 0)) = Add(.points(PointOf(1, iSpoke + 1)), v1)
        Next
        
        ' Then pentagons
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(2, iSpoke, 0)), .points(PointOf(1, iSpoke)))
            v2 = Subs(.points(PointOf(2, iSpoke, 1)), .points(PointOf(1, iSpoke)))
            v2 = Cross(v1, Cross(v1, v2))
            v1 = Scalar(v1, 1 / Mag(v1))
            v2 = Scalar(v2, -1 / Mag(v2))
            v3 = VectorOf(0, 0, 0)
            
            .points(PointOf(3, iSpoke, 1)) = Add(.points(PointOf(1, iSpoke)), Grid(VectorOf(50, 100 * (Sin(Rad(72)) + Cos(Rad(54))), 0), v1, v2, v3))
            .points(PointOf(3, iSpoke, 0)) = Add(.points(PointOf(1, iSpoke)), Grid(VectorOf(100 + 100 * Sin(Rad(18)), 100 * Sin(Rad(72)), 0), v1, v2, v3))
        Next
        
        ' Then hexagons again
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(3, iSpoke, 1)), .points(PointOf(2, iSpoke, 1)))
            v2 = Subs(.points(PointOf(3, iSpoke + 1, 0)), .points(PointOf(2, iSpoke + 1, 0)))
            
            .points(PointOf(4, iSpoke, 1)) = Add(.points(PointOf(3, iSpoke, 1)), v2)
            .points(PointOf(4, iSpoke + 1, 0)) = Add(.points(PointOf(3, iSpoke + 1, 0)), v1)
        Next
        
        ' Then hexagons again !
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(4, iSpoke, 0)), .points(PointOf(3, iSpoke, 0)))
            v2 = Subs(.points(PointOf(4, iSpoke, 1)), .points(PointOf(3, iSpoke, 1)))
            
            .points(PointOf(5, iSpoke, 0)) = Add(.points(PointOf(4, iSpoke, 0)), v2)
            .points(PointOf(5, iSpoke, 1)) = Add(.points(PointOf(4, iSpoke, 1)), v1)
        Next
        
        ' Then pentagons
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(5, iSpoke, 1)), .points(PointOf(4, iSpoke, 1)))
            v2 = Subs(.points(PointOf(4, iSpoke + 1, 0)), .points(PointOf(4, iSpoke, 1)))
            v2 = Cross(v1, Cross(v1, v2))
            v1 = Scalar(v1, 1 / Mag(v1))
            v2 = Scalar(v2, -1 / Mag(v2))
            v3 = VectorOf(0, 0, 0)
            
            '.points(PointOf(5, iSpoke + 1, 0)) = Add(.points(PointOf(4, iSpoke, 1)), Grid(VectorOf(50, 100 * (Sin(Rad(72)) + Cos(Rad(54))), 0), v1, v2, v3))
            .points(PointOf(6, iSpoke)) = Add(.points(PointOf(4, iSpoke, 1)), Grid(VectorOf(100 + 100 * Sin(Rad(18)), 100 * Sin(Rad(72)), 0), v1, v2, v3))
        Next
        
        ' More hexagons !
        For iSpoke = 0 To 4
            v1 = Subs(.points(PointOf(6, iSpoke - 1)), .points(PointOf(5, iSpoke, 0)))
            'v2 = Subs(.points(PointOf(4, iSpoke, 1)), .points(PointOf(3, iSpoke, 1)))
            
            .points(PointOf(7, iSpoke)) = Add(.points(PointOf(6, iSpoke)), v1)
            '.points(PointOf(5, iSpoke, 1)) = Add(.points(PointOf(4, iSpoke, 1)), v1)
        Next
        
        ' Join up lines
        For iSpoke = 0 To 4
            .lines(LineOf(0, iSpoke)) = JoinOf(PointOf(0, iSpoke), PointOf(0, iSpoke + 1))
            .lines(LineOf(1, iSpoke)) = JoinOf(PointOf(1, iSpoke), PointOf(0, iSpoke))
            .lines(LineOf(2, iSpoke, 0)) = JoinOf(PointOf(2, iSpoke, 0), PointOf(1, iSpoke))
            .lines(LineOf(2, iSpoke, 1)) = JoinOf(PointOf(2, iSpoke, 1), PointOf(1, iSpoke))
            .lines(LineOf(3, iSpoke)) = JoinOf(PointOf(2, iSpoke, 1), PointOf(2, iSpoke + 1, 0))
            .lines(LineOf(4, iSpoke, 0)) = JoinOf(PointOf(2, iSpoke, 0), PointOf(3, iSpoke, 0))
            .lines(LineOf(4, iSpoke, 1)) = JoinOf(PointOf(2, iSpoke, 1), PointOf(3, iSpoke, 1))
            .lines(LineOf(5, iSpoke)) = JoinOf(PointOf(3, iSpoke, 0), PointOf(3, iSpoke, 1))
            
            .lines(LineOf(6, iSpoke, 1)) = JoinOf(PointOf(3, iSpoke, 1), PointOf(4, iSpoke, 1))
            .lines(LineOf(6, iSpoke, 0)) = JoinOf(PointOf(3, iSpoke, 0), PointOf(4, iSpoke, 0))
            .lines(LineOf(7, iSpoke)) = JoinOf(PointOf(4, iSpoke, 1), PointOf(4, iSpoke + 1, 0))
            
            .lines(LineOf(8, iSpoke, 0)) = JoinOf(PointOf(5, iSpoke, 0), PointOf(4, iSpoke, 0))
            .lines(LineOf(8, iSpoke, 1)) = JoinOf(PointOf(5, iSpoke, 1), PointOf(4, iSpoke, 1))
            .lines(LineOf(9, iSpoke)) = JoinOf(PointOf(5, iSpoke, 0), PointOf(5, iSpoke, 1))
            
            .lines(LineOf(10, iSpoke, 1)) = JoinOf(PointOf(6, iSpoke), PointOf(5, iSpoke, 1))
            .lines(LineOf(10, iSpoke + 1, 0)) = JoinOf(PointOf(6, iSpoke), PointOf(5, iSpoke + 1, 0))
            
            .lines(LineOf(11, iSpoke)) = JoinOf(PointOf(7, iSpoke), PointOf(6, iSpoke))
            .lines(LineOf(12, iSpoke)) = JoinOf(PointOf(7, iSpoke - 1), PointOf(7, iSpoke))
        Next
        
        ' Create faces
        pentface.lineref(0) = LineOf(0, 0)
        pentface.lineref(1) = LineOf(0, 1)
        pentface.lineref(2) = LineOf(0, 2)
        pentface.lineref(3) = LineOf(0, 3)
        pentface.lineref(4) = LineOf(0, 4)
        .facets(0) = pentface
        
        For iSpoke = 0 To 4
            hexface.lineref(0) = LineOf(0, iSpoke)
            hexface.lineref(1) = LineOf(1, iSpoke + 1)
            hexface.lineref(2) = LineOf(2, iSpoke + 1, 0)
            hexface.lineref(3) = LineOf(3, iSpoke)
            hexface.lineref(4) = LineOf(2, iSpoke, 1)
            hexface.lineref(5) = LineOf(1, iSpoke)
            .facets(iSpoke + 1) = hexface
        Next
        
        For iSpoke = 0 To 4
            pentface.lineref(0) = LineOf(2, iSpoke, 0)
            pentface.lineref(1) = LineOf(2, iSpoke, 1)
            pentface.lineref(2) = LineOf(4, iSpoke, 1)
            pentface.lineref(3) = LineOf(5, iSpoke)
            pentface.lineref(4) = LineOf(4, iSpoke, 0)
            .facets(iSpoke + 6) = pentface
        Next

        For iSpoke = 0 To 4
            hexface.lineref(0) = LineOf(4, iSpoke, 1)
            hexface.lineref(1) = LineOf(3, iSpoke)
            hexface.lineref(2) = LineOf(4, iSpoke + 1, 0)
            hexface.lineref(3) = LineOf(6, iSpoke + 1, 0)
            hexface.lineref(4) = LineOf(7, iSpoke)
            hexface.lineref(5) = LineOf(6, iSpoke, 1)
            .facets(iSpoke + 11) = hexface
        Next
        
        For iSpoke = 0 To 4
            hexface.lineref(0) = LineOf(6, iSpoke, 0)
            hexface.lineref(1) = LineOf(5, iSpoke)
            hexface.lineref(2) = LineOf(6, iSpoke, 1)
            hexface.lineref(3) = LineOf(8, iSpoke, 1)
            hexface.lineref(4) = LineOf(9, iSpoke)
            hexface.lineref(5) = LineOf(8, iSpoke, 0)
            .facets(iSpoke + 16) = hexface
        Next
        
        For iSpoke = 0 To 4
            pentface.lineref(0) = LineOf(8, iSpoke, 1)
            pentface.lineref(1) = LineOf(10, iSpoke, 1)
            pentface.lineref(2) = LineOf(10, iSpoke + 1, 0)
            pentface.lineref(3) = LineOf(8, iSpoke + 1, 0)
            pentface.lineref(4) = LineOf(7, iSpoke)
            .facets(iSpoke + 21) = pentface
        Next
        
        For iSpoke = 0 To 4
            hexface.lineref(0) = LineOf(12, iSpoke)
            hexface.lineref(1) = LineOf(11, iSpoke)
            hexface.lineref(2) = LineOf(10, iSpoke, 1)
            hexface.lineref(3) = LineOf(9, iSpoke)
            hexface.lineref(4) = LineOf(10, iSpoke, 0)
            hexface.lineref(5) = LineOf(11, iSpoke - 1)
            .facets(iSpoke + 26) = hexface
        Next
        
        pentface.lineref(0) = LineOf(12, 0)
        pentface.lineref(4) = LineOf(12, 1)
        pentface.lineref(3) = LineOf(12, 2)
        pentface.lineref(2) = LineOf(12, 3)
        pentface.lineref(1) = LineOf(12, 4)
        .facets(31) = pentface

    
        ' Re-centre it
        Dim centrev As Vector
        
        For p = 0 To UBound(.points)
            centrev = Add(centrev, .points(p))
        Next
        centrev = Scalar(centrev, 1 / p)
        
        For p = 0 To UBound(.points)
            .points(p) = Subs(.points(p), centrev)
        Next
        
    End With
End Sub

Private Function Rad(degs As Integer) As Double
    Rad = degs * Atn(1) * 8 / 360
End Function

Private Function ModAdd(num As Integer, offset As Integer, mini As Integer, range As Integer, stepd As Integer) As Integer
    num = num + offset - mini
    While num < 0
        num = num + range * stepd
    Wend
    num = num Mod (range * stepd)
    num = num + mini
    ModAdd = num
End Function

Private Sub InitialiseCube()
    Cube.positional.axial.y = 1
    Cube.positional.position.z = 0
    Cube.positional.rightside.x = 1
    Cube.positional.down.z = 1
    Cube.movement.position = VectorOf(0, 0, 0)
    ReDim Cube.points(0 To 7) As Vector
    ReDim Cube.points2d(0 To 7) As Vector2d
    
    Cube.points(0) = VectorOf(-100, -100, -100)
    Cube.points(1) = VectorOf(100, -100, -100)
    Cube.points(2) = VectorOf(100, 100, -100)
    Cube.points(3) = VectorOf(-100, 100, -100)
    Cube.points(4) = VectorOf(-100, -100, 100)
    Cube.points(5) = VectorOf(100, -100, 100)
    Cube.points(6) = VectorOf(100, 100, 100)
    Cube.points(7) = VectorOf(-100, 100, 100)
    
    ReDim Cube.lines(0 To 11) As LineSet
    Cube.lines(0) = JoinOf(0, 1)
    Cube.lines(1) = JoinOf(1, 2)
    Cube.lines(2) = JoinOf(2, 3)
    Cube.lines(3) = JoinOf(3, 0)
    Cube.lines(4) = JoinOf(4, 5)
    Cube.lines(5) = JoinOf(5, 6)
    Cube.lines(6) = JoinOf(6, 7)
    Cube.lines(7) = JoinOf(7, 4)
    Cube.lines(8) = JoinOf(0, 4)
    Cube.lines(9) = JoinOf(1, 5)
    Cube.lines(10) = JoinOf(2, 6)
    Cube.lines(11) = JoinOf(3, 7)
    
    Dim vs As VectorSet
    
    ReDim Cube.facets(0 To 5) As VectorSet
    ReDim vs.lineref(0 To 3) As Integer
    
    vs.lineref(0) = 3
    vs.lineref(1) = 2
    vs.lineref(2) = 1
    vs.lineref(3) = 0
    
    Cube.facets(0) = vs
    
    vs.lineref(0) = 4
    vs.lineref(1) = 5
    vs.lineref(2) = 6
    vs.lineref(3) = 7
    Cube.facets(1) = vs
    
    vs.lineref(0) = 0
    vs.lineref(1) = 9
    vs.lineref(2) = 4
    vs.lineref(3) = 8
    Cube.facets(2) = vs
    
    vs.lineref(0) = 1
    vs.lineref(1) = 10
    vs.lineref(2) = 5
    vs.lineref(3) = 9
    Cube.facets(3) = vs
    
    vs.lineref(0) = 2
    vs.lineref(1) = 11
    vs.lineref(2) = 6
    vs.lineref(3) = 10
    Cube.facets(4) = vs
    
    vs.lineref(0) = 3
    vs.lineref(1) = 8
    vs.lineref(2) = 7
    vs.lineref(3) = 11
    Cube.facets(5) = vs
End Sub

Private Sub InitialiseViewPoint()
    MyAxis = VectorOf(0, 0, 1)
    MyRightSide = VectorOf(1, 0, 0)
    RollRate = 0
    PitchRate = 0
    RollAbs = 0
    PitchAbs = 0
End Sub

Private Sub Form_Activate()
    InitialiseViewPoint
    'InitialiseCube
    InitialiseFootball
End Sub

Private Sub Render(thing As Entity, forecolour As Long, hiddencolour As Long)
    Dim Join As LineSet
    Dim startVector As Vector2d
    Dim endVector As Vector2d
    
    Dim p As Integer
    Dim q As Integer
    
    Dim va As Vector2d
    Dim vb As Vector2d
    Dim vis As Boolean
    
    Dim pos As Vector
    
    If thing.positional.position.z > 0 Then
        Exit Sub
    End If
    
    For p = LBound(thing.points) To UBound(thing.points)
        ' My positional offset
        ' My absolute axis
        pos = Add(thing.positional.position, Grid(thing.points(p), thing.positional.rightside, thing.positional.axial, thing.positional.down))
        'Rotate pos, MyAngle, MyAxis
        thing.points2d(p) = ConvertTo2d(pos)
    Next
    
    For p = LBound(thing.lines) To UBound(thing.lines)
        thing.lines(p).visible = False
    Next
    
    For p = LBound(thing.facets) To UBound(thing.facets)
        Join = thing.lines((thing.facets(p).lineref(0)))
        va = Sub2d(thing.points2d(Join.finish), thing.points2d(Join.start))

        Join = thing.lines(thing.facets(p).lineref(1))
        vb = Sub2d(thing.points2d(Join.finish), thing.points2d(Join.start))
        
        vis = Cross2d(va, vb) < 0
        For q = LBound(thing.facets(p).lineref) To UBound(thing.facets(p).lineref)
            thing.lines(thing.facets(p).lineref(q)).visible = thing.lines(thing.facets(p).lineref(q)).visible Or vis
        Next
    Next
    
    For p = LBound(thing.lines) To UBound(thing.lines)
        If thing.lines(p).visible Then
            Join = thing.lines(p)
            startVector = thing.points2d(Join.start)
            endVector = thing.points2d(Join.finish)
            Me.Line (500 + startVector.x + mewidth, 500 + startVector.y + meheight)-(500 + endVector.x + mewidth, 500 + endVector.y + meheight), forecolour
        Else
'            Join = thing.lines(p)
'            startVector = thing.points2d(Join.start)
'            endVector = thing.points2d(Join.finish)
'            Me.Line (500 + startVector.x + mewidth, 500 + startVector.y + meheight)-(500 + endVector.x + mewidth, 500 + endVector.y + meheight), hiddencolour
        End If
    Next
End Sub


Private Sub Render2(thing As Entity)
    Dim Join As LineSet
    Dim startVector As Vector2d
    Dim endVector As Vector2d
    
    Dim p As Integer
    Dim q As Integer
    
    Dim va As Vector2d
    Dim vb As Vector2d
    Dim va3d As Vector
    Dim vb3d As Vector
    
    Dim vis As Boolean
    
    Dim pos As Vector
    Dim pos1 As Vector
    Dim pos2 As Vector
    
    Dim shade As Double
    
    If thing.positional.position.z > 0 Then
        Exit Sub
    End If
    
    For p = LBound(thing.points) To UBound(thing.points)
        ' My positional offset
        ' My absolute axis
        pos = Add(thing.positional.position, Grid(thing.points(p), thing.positional.rightside, thing.positional.axial, thing.positional.down))
        'Rotate pos, MyAngle, MyAxis
        thing.points2d(p) = ConvertTo2d(pos)
    Next
    
    For p = LBound(thing.facets) To UBound(thing.facets)
        ' Visible ?
        Join = thing.lines((thing.facets(p).lineref(0)))
        va = Sub2d(thing.points2d(Join.finish), thing.points2d(Join.start))

        Join = thing.lines(thing.facets(p).lineref(1))
        vb = Sub2d(thing.points2d(Join.finish), thing.points2d(Join.start))
        
        vis = Cross2d(va, vb) < 0
        thing.facets(p).visible = vis
        
        ' Shade?
        If vis Then
            Join = thing.lines((thing.facets(p).lineref(0)))
            pos1 = Add(thing.positional.position, Grid(thing.points(Join.start), thing.positional.rightside, thing.positional.axial, thing.positional.down))
            pos2 = Add(thing.positional.position, Grid(thing.points(Join.finish), thing.positional.rightside, thing.positional.axial, thing.positional.down))
            'va3d = Subs(thing.points(Join.finish), thing.points(Join.start))
            va3d = Subs(pos2, pos1)
    
            Join = thing.lines(thing.facets(p).lineref(1))
            pos1 = Add(thing.positional.position, Grid(thing.points(Join.start), thing.positional.rightside, thing.positional.axial, thing.positional.down))
            pos2 = Add(thing.positional.position, Grid(thing.points(Join.finish), thing.positional.rightside, thing.positional.axial, thing.positional.down))
            'vb3d = Subs(thing.points(Join.finish), thing.points(Join.start))
            vb3d = Subs(pos2, pos1)
            
            Dim normal As Vector
            Dim source As Vector
            
            normal = Cross(va3d, vb3d)
            normal = Scalar(normal, 1 / Mag(normal))
            source = VectorOf(1000, 1000, 0)
            source = Scalar(source, 1 / Mag(source))
            
            shade = 100 + 128 * (1 - Dot(normal, source))
                            
            If UBound(thing.facets(p).lineref) = 4 Then
                shade = shade / 2
            End If
            Drawface thing, p, RGB(shade, shade, shade)
            
        End If
    Next
    
End Sub

Private Sub Drawface(othing As Entity, faceindex As Integer, ByVal lColour As Long)
    Dim hdc As Long
    Dim result As Long
    Dim i As Integer
    Dim num As Integer
    Dim pixel() As POINTAPI
    
    Dim sortpoints() As Integer
    Dim marked() As Boolean
    Dim iCurrent As Integer
    
    hdc = Me.hdc
    num = UBound(othing.facets(faceindex).lineref)
    ReDim pixel(0 To num) As POINTAPI
    
    ReDim marked(0 To num) As Boolean
    ReDim sortpoints(0 To num + 1) As Integer
    Dim sortpointsindex As Integer
    Dim bFinished As Boolean
    Dim pointcheck As Integer
    
    With othing
        sortpoints(0) = .lines(.facets(faceindex).lineref(0)).start
        sortpoints(1) = .lines(.facets(faceindex).lineref(0)).finish
        marked(0) = True
        sortpointsindex = 2
        iCurrent = .lines(.facets(faceindex).lineref(0)).finish
               
        While Not bFinished
            bFinished = True
            For pointcheck = 0 To num
                If Not marked(pointcheck) Then
                    If iCurrent = .lines(.facets(faceindex).lineref(pointcheck)).start Then
                        marked(pointcheck) = True
                        bFinished = False
                        iCurrent = .lines(.facets(faceindex).lineref(pointcheck)).finish
                        sortpoints(sortpointsindex) = iCurrent
                        sortpointsindex = sortpointsindex + 1
                    ElseIf iCurrent = .lines(.facets(faceindex).lineref(pointcheck)).finish Then
                        marked(pointcheck) = True
                        bFinished = False
                        iCurrent = .lines(.facets(faceindex).lineref(pointcheck)).start
                        sortpoints(sortpointsindex) = iCurrent
                        sortpointsindex = sortpointsindex + 1
                    End If
                End If
            Next
        Wend
    End With
    
    For i = 0 To num
        pixel(i).x = ScaleX(3500 + othing.points2d(sortpoints(i)).x, 1, 3)
        pixel(i).y = ScaleY(3500 + othing.points2d(sortpoints(i)).y, 1, 3)
    Next
    
    Dim rgn As Long
    Dim brush As Long
    Dim bs As LOGBRUSH
    
    bs.lbColor = lColour
    bs.lbHatch = 0
    bs.lbStyle = BS_SOLID
    
    num = num + 1
    brush = CreateBrushIndirect(bs)
    rgn = CreatePolygonRgn(pixel(0), num, 1)
    result = FillRgn(hdc, rgn, brush)
    DeleteObject brush
    DeleteObject rgn
    
End Sub

Private Sub Update(thing As Entity)

    thing.positional.position = Add(thing.positional.position, thing.movement.position)
     
    Rotate thing.positional.down, roll, thing.positional.axial
    Rotate thing.positional.rightside, roll, thing.positional.axial
    
    Rotate thing.positional.axial, pitch, thing.positional.rightside
    Rotate thing.positional.down, pitch, thing.positional.rightside
    
    'thing.positional.position = Add(thing.positional.position, Scalar(myAxis, Speed))
    'RollAbs = RollAbs + RollRate
    'PitchAbs = PitchAbs + PitchRate
End Sub

Private Sub Form_Load()
    pi = Atn(1) * 4
    roll = -pi / 500
    pitch = pi / 500
    mewidth = Me.Width / 2
    meheight = Me.Height / 2
End Sub

Private Sub Timer1_Timer()
'    Render Football, &H7C7303, &H7C7303
'    Update Football
'    MyAngle = MyAngle + 5 * pi / 360
' '   Render Football, vbCyan, RGB(0, 164, 164)
'    Render Football, vbCyan, &H7C7303

    Me.Cls
    Render2 Football
    Update Football
    MyAngle = MyAngle + 5 * pi / 360
End Sub

' Calculate point number according to criteria
Private Function PointOf(iLevel As Integer, iSpoke As Integer, Optional iLeftRight As Integer = 0) As Integer
    While iSpoke < 0
        iSpoke = iSpoke + 5
    Wend
    iSpoke = iSpoke Mod 5
    
    Select Case iLevel
        Case 0
            PointOf = iSpoke
        Case 1
            PointOf = iSpoke + 5
        Case 2
            PointOf = iSpoke * 2 + 10 + iLeftRight
        Case 3
            PointOf = iSpoke * 2 + 20 + iLeftRight
        Case 4
            PointOf = iSpoke * 2 + 30 + iLeftRight
        Case 5
            PointOf = iSpoke * 2 + 40 + iLeftRight
        Case 6
            PointOf = iSpoke + 50
        Case 7
            PointOf = iSpoke + 55
    End Select

End Function

Private Function LineOf(iLevel As Integer, iSpoke As Integer, Optional iLeftRight As Integer = 0) As Integer
    While iSpoke < 0
        iSpoke = iSpoke + 5
    Wend
    iSpoke = iSpoke Mod 5
    
    Select Case iLevel
        Case 0
            LineOf = iSpoke
        Case 1
            LineOf = iSpoke + 5
        Case 2
            LineOf = iSpoke * 2 + 10 + iLeftRight
        Case 3
            LineOf = iSpoke + 20
        Case 4
            LineOf = iSpoke * 2 + 25 + iLeftRight
        Case 5
            LineOf = iSpoke + 35
        Case 6
            LineOf = iSpoke * 2 + 40 + iLeftRight
        Case 7
            LineOf = iSpoke + 50
        Case 8
            LineOf = iSpoke * 2 + 55 + iLeftRight
        Case 9
            LineOf = iSpoke + 65
        Case 10
            LineOf = iSpoke * 2 + 70 + iLeftRight
        Case 11
            LineOf = iSpoke + 80
        Case 12
            LineOf = iSpoke + 85
    End Select

End Function
