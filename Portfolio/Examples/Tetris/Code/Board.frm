VERSION 5.00
Begin VB.Form Board 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tetris"
   ClientHeight    =   5970
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3885
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080C0FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Ticker 
      Interval        =   10
      Left            =   840
      Top             =   1200
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private running As Boolean


Private World(-4 To 15, -4 To 25) As Boolean
Private dirx(0 To 4) As Variant

Private shapes(1 To 7) As String
Private currentshape As Integer
Private currentsx As Integer
Private currentsy As Integer
Private currentorient As Integer

Private colide As Boolean
Private leftwall As Boolean
Private rightwall As Boolean

Private speed As Integer
Private counter As Integer
Private downdisabled As Boolean

Private Const KeyLeft = 37
Private Const KeyRight = 39
Private Const KeyDown = 40
Private Const KeyAnticlock = 90
Private Const KeyClock = 88

Private iLineCount As Integer
Private updating As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KeyLeft
            MoveShape -1, 0, 0
        Case KeyRight
            MoveShape 1, 0, 0
        Case KeyDown
            If Not downdisabled Then
                If MoveShape(0, -1, 0) Then
                    counter = speed - 2
                    downdisabled = True
                End If
            End If
        Case KeyAnticlock
            MoveShape 0, 0, -1
        Case KeyClock
            MoveShape 0, 0, 1
    End Select
End Sub

Private Sub Form_Load()
    shapes(1) = "SUSDDSDS"
    shapes(2) = "SLSRRSDS"
    shapes(3) = "SLSRRSUS"
    shapes(4) = "SRSDSLS"
    shapes(5) = "SLSRUSDRS"
    shapes(6) = "SLSRUSRS"
    shapes(7) = "SLSRDSRS"
    
    dirx(0) = Array(1, 0) ' right
    dirx(1) = Array(0, -1) ' down
    dirx(2) = Array(-1, 0) ' left
    dirx(3) = Array(0, 1) ' up
    

    Dim x As Integer
    Dim y As Integer
    
    For x = 0 To UBound(World, 1)
        World(x, 0) = True
    Next
    For y = 0 To UBound(World, 2)
        World(0, y) = True
        World(UBound(World, 1), y) = True
    Next
    
    speed = 100
    iLineCount = 0
    running = True
    
    ' Draw holder
    Me.Line (16 * Screen.TwipsPerPixelX - 5 * 16, Me.ScaleHeight - 10 * Screen.TwipsPerPixelY)-Step(14 * 16 * Screen.TwipsPerPixelX + 10 * 16, 0)
    Me.Line (16 * Screen.TwipsPerPixelX - 5 * 16, Me.ScaleHeight - 10 * Screen.TwipsPerPixelY)-Step(0, -22 * 16 * Screen.TwipsPerPixelY)
    Me.Line (15 * 16 * Screen.TwipsPerPixelX + 5 * 16, Me.ScaleHeight - 10 * Screen.TwipsPerPixelY)-Step(0, -22 * 16 * Screen.TwipsPerPixelY)
    
    ShowLines
    NewShape
End Sub

Private Function ShowLines()
    Me.DrawMode = vbCopyPen
    Me.Line (16, 16)-Step(Me.ScaleWidth, 32 * Screen.TwipsPerPixelY), vbBlack, BF
    Me.Line (0, 0)-Step(0, 0)
    Me.Print "LINES " & iLineCount
    Me.DrawMode = vbXorPen
    
    If iLineCount < 100 Then
        speed = 60 * (7 / 60) ^ (iLineCount / 120)
    Else
        speed = 1
    End If
End Function

Private Function MoveShape(xdir As Integer, ydir As Integer, rotate As Integer) As Boolean
    Dim leftwall As Boolean
    Dim rightwall As Boolean
    Dim collided As Boolean
    
    DrawShape currentshape, currentsx, currentsy, currentorient, collided
    
    currentsx = currentsx + xdir
    currentsy = currentsy + ydir
    currentorient = (currentorient + 4 + rotate) Mod 4
    
    DrawShape currentshape, currentsx, currentsy, currentorient, collided
    If collided Then
        MoveShape = True
        DrawShape currentshape, currentsx, currentsy, currentorient, collided
        currentsx = currentsx - xdir
        currentsy = currentsy - ydir
        currentorient = (currentorient + 4 - rotate) Mod 4
        DrawShape currentshape, currentsx, currentsy, currentorient, collided
    End If
End Function

Private Sub DrawShape(ByVal iShape As Integer, ByVal xpos As Integer, ByVal ypos As Integer, ByVal orientation As Integer, Optional collided As Boolean, Optional ByVal fill As Boolean)

    Dim ins As String
    Dim iIndex As Integer

    For iIndex = 1 To Len(shapes(iShape))
        Select Case Mid$(shapes(iShape), iIndex, 1)
            Case "S"
                If Not fill Then
                    Me.Line (xpos * 16 * 15, Me.ScaleHeight - ypos * 16 * 15)-Step(16 * 15, -16 * 15), vbRed, BF
                    Me.Line (xpos * 16 * 15 + 15, Me.ScaleHeight - ypos * 16 * 15 - 15)-Step(16 * 15 - 30, -16 * 15 + 30), , BF
                    If World(xpos, ypos) Then
                        collided = True
                    End If
                Else
                    World(xpos, ypos) = True
                End If
            Case "L"
                xpos = xpos + dirx((2 + orientation) Mod 4)(0)
                ypos = ypos + dirx((2 + orientation) Mod 4)(1)
            Case "R"
                xpos = xpos + dirx((0 + orientation) Mod 4)(0)
                ypos = ypos + dirx((0 + orientation) Mod 4)(1)
            Case "U"
                xpos = xpos + dirx((3 + orientation) Mod 4)(0)
                ypos = ypos + dirx((3 + orientation) Mod 4)(1)
            Case "D"
                xpos = xpos + dirx((1 + orientation) Mod 4)(0)
                ypos = ypos + dirx((1 + orientation) Mod 4)(1)
        End Select
    Next
End Sub

Private Sub mnuHelp_Click()
    Ticker.Enabled = False
    frmHelp.Show vbModal
    Ticker.Enabled = True
End Sub

Private Sub Ticker_Timer()
    Dim dummy As Boolean
    
    counter = counter + 1
    If (counter Mod speed) = 0 Then
        If MoveShape(0, -1, 0) Then
            DrawShape currentshape, currentsx, currentsy, currentorient, dummy, True
            iLineCount = iLineCount + CheckRows
            ShowLines
            NewShape
            downdisabled = False
        End If
        counter = 0
    End If
End Sub

Private Sub NewShape()
    Dim terminate_game  As Boolean
    
    Randomize
    currentshape = 1 + Int(Rnd * 7)
    currentsx = UBound(World, 1) \ 2
    currentsy = UBound(World, 2) - 4
    currentorient = Rnd * 4
    
    DrawShape currentshape, currentsx, currentsy, currentorient, terminate_game
    If terminate_game Then
        Ticker.Enabled = False
        MsgBox "You completed " & iLineCount & " lines." & vbCrLf & vbCrLf & "Press OK for another game", vbOKOnly, "Game Over"
        Me.Cls
        Erase World
        Form_Load
        Ticker.Enabled = True
    End If
End Sub

Private Function CheckRows() As Integer
    Dim iRow As Integer
    Dim iColumn  As Integer
    Dim iAdd As Integer
    Dim iDelCol As Integer
    Dim iDelRow As Integer
    
    For iRow = UBound(World, 2) To 1 Step -1
        iAdd = 0
        For iColumn = 1 To UBound(World, 1) - 1
            If World(iColumn, iRow) Then
                iAdd = iAdd + 1
            End If
        Next
        If iAdd = UBound(World, 1) - 1 Then
            CheckRows = CheckRows + 1
'            For iDelCol = 1 To UBound(World, 1)
'                Block iDelCol, iRow
'            Next
            For iDelRow = iRow To UBound(World, 2) - 1
                For iDelCol = 1 To UBound(World, 1)
                    If World(iDelCol, iDelRow) Then
                        Block iDelCol, iDelRow
                    End If
                    World(iDelCol, iDelRow) = World(iDelCol, iDelRow + 1)
                    If World(iDelCol, iDelRow + 1) Then
                        Block iDelCol, iDelRow
                    End If
                Next
            Next
        End If
    Next
End Function

Private Sub Block(ByVal xpos As Integer, ByVal ypos As Integer)
    Me.Line (xpos * 16 * 15, Me.ScaleHeight - ypos * 16 * 15)-Step(16 * 15, -16 * 15), vbRed, BF
    Me.Line (xpos * 16 * 15 + 15, Me.ScaleHeight - ypos * 16 * 15 - 15)-Step(16 * 15 - 30, -16 * 15 + 30), , BF
End Sub



