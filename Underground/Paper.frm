VERSION 5.00
Begin VB.Form Paper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Relationships"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   11865
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Johnston ITC Std Medium"
      Size            =   20.25
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRGB 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   2
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   1
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   0
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private mnMouseX As Single
Private mnMouseY As Single

Private oSelectedPositions As Collection

Private oSelectedPosition As Position
Private oSelectedPosOffsetX As Single
Private oSelectedPosOffsetY As Single

Private nBoxStartX As Single
Private nBoxStartY As Single

Private oInitialLink As Position
Private oFinalLink As Position

Private mnPreviousX As Single
Private mnPreviousY As Single
    
Public BackColour As Long

Public DiagramRef As New Diagram

Private bDragGroupSelected As Boolean

Private mlRecentColourIndex As Long

Public mbGraticuleOn As Boolean
Public mbShowCircles As Boolean

Private Sub UnselectDragGroup()
    Dim oPosition As Position
    
    bDragGroupSelected = False
    If Not oSelectedPositions Is Nothing Then
        For Each oPosition In oSelectedPositions
            oPosition.RenderName
        Next
        Set oSelectedPositions = Nothing
    End If
       'DiagramRef.Render
End Sub

Private Sub Form_Activate()
    Dim c As New CirclePrimitive
    Dim oCentre As New Vector
    Dim ostart As New Vector
    Dim X As Long
    Dim Y As Long
    Dim lPixelOn As Long
    
    oCentre.SetVector 20, 20
    c.Initialise oCentre, 5
    c.SetStart ostart
    
    For Y = 0 To 40
        For X = 0 To 40
            If Y Mod 2 = 0 Then
                lPixelOn = lPixelOn Xor c.MoveRight
            Else
                lPixelOn = lPixelOn Xor c.MoveLeft
            End If
            
        Next
        lPixelOn = lPixelOn Xor c.MoveDown
    Next
End Sub

Private Sub Form_Initialize()
    Set DiagramRef.PaperRef = Me
    DiagramRef.Colours = Array(RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 255, 0), RGB(255, 255, 0), RGB(0, 0, 255), RGB(255, 0, 255), RGB(0, 255, 255), RGB(255, 255, 255), RGB(128, 128, 128), &H80FF&, 16576, &H800080, RGB(255, 128, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(0, 0, 0), RGB(255, 255, 255))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oPosition As Position
    Dim oPosition2 As Position
    Dim oRelationship As Relationship
    Dim vRelColours As Variant
    Dim lColourIndex As Long
    Dim lColour As Long
    Dim yColour(3) As Byte
    Dim vText As Variant
    Dim vLine As Variant
    Dim lIndex As Long
    Dim oLastPosition As Position
    Dim oCurvePositions(3) As Position
    
    Select Case KeyCode
   'Single Functions
        Case vbKeyV
            If Shift = 2 Then
                If Clipboard.GetFormat(vbCFText) Then
                    vText = Split(Clipboard.GetText(vbCFText), vbCrLf)
                    lIndex = 0
                    Set oLastPosition = Nothing
                    For Each vLine In vText
                        Set oPosition = New Position
                        oPosition.Name = vLine
                        oPosition.Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
                        oPosition.Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20 + lIndex * 20
                        oPosition.ColourIndex = mlRecentColourIndex
                        Set oPosition.DiagramRef = DiagramRef
                        Set oPosition.ParserRef = oParsePosition
                        DiagramRef.Positions.List.Add oPosition
                        
                        If Not oLastPosition Is Nothing Then
                            Set oRelationship = New Relationship
                            Set oRelationship.DiagramRef = DiagramRef
                            Set oRelationship.FromPos = oLastPosition
                            Set oRelationship.ToPos = oPosition
                            oRelationship.ColourIndeces = Array(mlRecentColourIndex)
                            oRelationship.RenderRelationship
                            DiagramRef.Relationships.List.Add oRelationship
                        End If
                        Set oLastPosition = oPosition
                        lIndex = lIndex + 1
                    Next
                    DiagramRef.RenderAndSave
                End If
                UnselectDragGroup
            End If
        Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKey0
            mlRecentColourIndex = (KeyCode - 48 - ((Shift And 1) = 1) * 10)

            lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            If oInitialLink Is Nothing Then
                If (Shift And 2) = 2 Then
                    If Not oSelectedPositions Is Nothing Then
                        For Each oPosition In oSelectedPositions
                            oPosition.ColourIndex = KeyCode - 48 - ((Shift And 1) = 1) * 10
                            oPosition.RenderName
                            For Each oPosition2 In oSelectedPositions
                                If Not oPosition2 Is oPosition Then
                                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oPosition, oPosition2)
                                    If Not oRelationship Is Nothing Then
                                        oRelationship.ColourIndeces = Array(KeyCode - 48 - ((Shift And 1) = 1) * 10)
                                    End If
                                End If
                            Next
                        Next
                        DiagramRef.RenderAndSave
                    Else
                        Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                        If Not oPosition Is Nothing Then
                            oPosition.ColourIndex = KeyCode - 48 - ((Shift And 1) = 1) * 10
                            oPosition.RenderName
                        End If
                    End If
                Else
                    Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                    If oInitialLink Is Nothing Then
                        NewString.txtString = ""
                        NewString.Show vbModal
                        Set oPosition = New Position
                        oPosition.Name = NewString.txtString
                        oPosition.Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
                        oPosition.Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
                        oPosition.ColourIndex = (KeyCode - 48 - (Shift = 1) * 10)
                        Set oPosition.DiagramRef = DiagramRef
                        Set oPosition.ParserRef = oParsePosition
                        DiagramRef.Positions.List.Add oPosition
                        oPosition.RenderName
                        DiagramRef.FileIOs.WriteFile
                    Else
                        oInitialLink.RenderName &HFFC0C0
                    End If
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    If oFinalLink.Reference = oInitialLink.Reference Then
                        oInitialLink.RenderName
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                Else
                    oInitialLink.RenderName
                    Set oInitialLink = Nothing
                    Set oFinalLink = Nothing
                End If
                
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If oRelationship Is Nothing Then
                        Set oRelationship = New Relationship
                        Set oRelationship.DiagramRef = DiagramRef
                        Set oRelationship.FromPos = oInitialLink
                        Set oRelationship.ToPos = oFinalLink
                        oRelationship.ColourIndeces = Array((KeyCode - 48 - (Shift = 1) * 10))
                        oRelationship.RenderRelationship
                        DiagramRef.Relationships.List.Add oRelationship
                        DiagramRef.Positions.RenderAll
                    Else
                        lColourIndex = (KeyCode - 48 - (Shift = 1) * 10)
                        vRelColours = oRelationship.ColourIndeces
                        If InArray(vRelColours, lColourIndex) Then
                            RemoveFromArray vRelColours, lColourIndex
                        Else
                            AddToArray vRelColours, lColourIndex
                        End If
                        oRelationship.ColourIndeces = vRelColours
                        If UBound(vRelColours) = -1 Then
                            DiagramRef.Relationships.RemoveRelationship oRelationship
                        End If
                        DiagramRef.RenderAndSave
                    End If
                    Set oInitialLink = Nothing
                    Set oFinalLink = Nothing
                    DiagramRef.FileIOs.WriteFile
                End If
            End If
            UnselectDragGroup
        Case vbKeyB
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If Not oRelationship Is Nothing Then
                        oRelationship.Style = 1 - oRelationship.Style
                        DiagramRef.RenderAndSave
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                End If
            End If
            UnselectDragGroup
        Case vbKeyD
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If Not oRelationship Is Nothing Then
                        DiagramRef.Relationships.RemoveRelationship oRelationship
                        DiagramRef.RenderAndSave
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                End If
            End If
            UnselectDragGroup
        Case vbKeyL
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If Not oRelationship Is Nothing Then
                        oRelationship.Style = 2
                        DiagramRef.RenderAndSave
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                End If
            End If
            UnselectDragGroup
        Case vbKeyE
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                oPosition.ClearName
                NewString.txtString = oPosition.Name
                NewString.Show vbModal
                oPosition.Name = NewString.txtString
                oPosition.RenderName
                DiagramRef.FileIOs.WriteFile
                UnselectDragGroup
            End If
        Case vbKeyG
            mbGraticuleOn = Not mbGraticuleOn
            DiagramRef.RenderAndSave
        Case vbKeyAdd
            DiagramRef.Zoom = DiagramRef.Zoom * 4 / 3
            DiagramRef.RenderAndSave
        Case vbKeySubtract
            DiagramRef.Zoom = Int(DiagramRef.Zoom * 3 + 0.5) / 4
            DiagramRef.RenderAndSave
        Case vbKeyS
            mbShowCircles = Not mbShowCircles
            DiagramRef.Relationships.RemoveDuplicates
            DiagramRef.RenderAndSave
        Case vbKeyF
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                If Shift <> 2 Then
                    DiagramRef.Positions.SendToFront oPosition
                    DiagramRef.Relationships.SendToFront oPosition
                Else
                    DiagramRef.Positions.SendToBack oPosition
                    DiagramRef.Relationships.SendToBack oPosition
                End If
                DiagramRef.RenderAndSave
                UnselectDragGroup
            End If
        Case vbKeyT
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    If Shift <> 2 Then
                        DiagramRef.Positions.SendToFront oInitialLink
                        DiagramRef.Positions.SendToFront oFinalLink
                        DiagramRef.Relationships.SendToFront oInitialLink, oFinalLink
                    Else
                        DiagramRef.Positions.SendToBack oInitialLink
                        DiagramRef.Positions.SendToBack oFinalLink
                        DiagramRef.Relationships.SendToBack oInitialLink, oFinalLink
                    End If
                    DiagramRef.RenderAndSave
                End If
                Set oInitialLink = Nothing
                Set oFinalLink = Nothing
            End If
            UnselectDragGroup
        Case vbKeyC
            scrColourComponent(0).Visible = Not scrColourComponent(0).Visible
            scrColourComponent(1).Visible = Not scrColourComponent(1).Visible
            scrColourComponent(2).Visible = Not scrColourComponent(2).Visible
            txtRGB.Visible = Not txtRGB.Visible
        Case vbKeyI
            If Not oSelectedPositions Is Nothing Then
                Dim nAverage As Single
                Dim nTotal As Single
                
                nTotal = 0
                For Each oPosition In oSelectedPositions
                    nAverage = nAverage + oPosition.Pos.Y
                    nTotal = nTotal + 1
                Next
                nAverage = nAverage / nTotal
                For Each oPosition In oSelectedPositions
                    oPosition.Pos.Y = 2 * nAverage - oPosition.Pos.Y
                Next
                DiagramRef.RenderAndSave
            End If
        Case vbKeyR
            If (Shift And 2) = 2 Then
                If oSelectedPositions Is Nothing Then
                    Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                    If Not oPosition Is Nothing Then
                        oPosition.Proximity = (oPosition.Proximity + (Shift And 1) * -2 + 1)
                        DiagramRef.RenderAndSave
                    End If
                Else
                    For Each oPosition In oSelectedPositions
                        oPosition.Proximity = (oPosition.Proximity + (Shift And 1) * -2 + 1)
                    Next
                    DiagramRef.RenderAndSave
                End If
            Else
                If oSelectedPositions Is Nothing Then
                    Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                    If Not oPosition Is Nothing Then
                        oPosition.Orientation = (oPosition.Orientation + 2 * (Shift = 1) + 1 + 8) Mod 8
                        DiagramRef.RenderAndSave
                    End If
                Else
                    For Each oPosition In oSelectedPositions
                        oPosition.Orientation = (oPosition.Orientation + 2 * (Shift = 1) + 1 + 8) Mod 8
                    Next
                    DiagramRef.RenderAndSave
                End If
            End If
        Case vbKeyP
            If oSelectedPositions Is Nothing Then
                Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oPosition Is Nothing Then
                    oPosition.Shape = (oPosition.Shape + 2 * (Shift = 1) + 1 + 8) Mod 8
                    DiagramRef.RenderAndSave
                End If
            Else
                For Each oPosition In oSelectedPositions
                    oPosition.Shape = (oPosition.Shape + 2 * (Shift = 1) + 1 + 8) Mod 8
                Next
                DiagramRef.RenderAndSave
            End If
        Case vbKeyDelete
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                UnselectDragGroup
                DiagramRef.Relationships.RemoveRelationshipWithReference oPosition
                DiagramRef.Positions.RemovePosition oPosition
                DiagramRef.RenderAndSave
            Else
                If Not oSelectedPositions Is Nothing Then
                    For Each oPosition In oSelectedPositions
                        DiagramRef.Relationships.RemoveRelationshipWithReference oPosition
                        DiagramRef.Positions.RemovePosition oPosition
                    Next
                    DiagramRef.RenderAndSave
                    Set oSelectedPositions = Nothing
                End If
            End If
        Case vbKeyK
            Dim fXAv As Double
            Dim fYAv As Double
            
            For Each oPosition In DiagramRef.Positions.List
                fXAv = fXAv + oPosition.Pos.X
                fYAv = fYAv + oPosition.Pos.Y
            Next
            If DiagramRef.Positions.List.Count > 0 Then
                DiagramRef.TopLeft.X = Me.Width / Screen.TwipsPerPixelX / 2 - fXAv / DiagramRef.Positions.List.Count
                DiagramRef.TopLeft.Y = Me.Height / Screen.TwipsPerPixelY / 2 - fYAv / DiagramRef.Positions.List.Count
            Else
                DiagramRef.TopLeft.X = Me.Width / Screen.TwipsPerPixelX / 2
                DiagramRef.TopLeft.Y = Me.Height / Screen.TwipsPerPixelY / 2
            End If
            DiagramRef.RenderAndSave
        Case vbKeySpace
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                If oPosition.Name = " " Then
                    oPosition.Name = ""
                Else
                    oPosition.Name = " "
                End If
                DiagramRef.RenderAndSave
            End If
        Case vbKeyLeft
            lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            For lIndex = 0 To 3
                Set oCurvePositions(lIndex) = New Position
                
                With oCurvePositions(lIndex)
                    Set .DiagramRef = DiagramRef
                    Set .ParserRef = oParsePosition
                    DiagramRef.Positions.List.Add oCurvePositions(lIndex)
                End With
            Next
            oCurvePositions(0).Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
            oCurvePositions(0).Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
            
            oCurvePositions(1).Pos.X = oCurvePositions(0).Pos.X - 20
            oCurvePositions(1).Pos.Y = oCurvePositions(0).Pos.Y
            
            oCurvePositions(2).Pos.X = oCurvePositions(1).Pos.X - 20
            oCurvePositions(2).Pos.Y = oCurvePositions(1).Pos.Y + 20
            
            oCurvePositions(3).Pos.X = oCurvePositions(2).Pos.X
            oCurvePositions(3).Pos.Y = oCurvePositions(2).Pos.Y + 20

            For lIndex = 0 To 2
                Set oRelationship = New Relationship
                Set oRelationship.DiagramRef = DiagramRef
                Set oRelationship.FromPos = oCurvePositions(lIndex)
                Set oRelationship.ToPos = oCurvePositions(lIndex + 1)
                If lIndex = 1 Then
                    oRelationship.Style = 1
                End If
                oRelationship.ColourIndeces = Array(mlRecentColourIndex)
                oRelationship.RenderRelationship
                DiagramRef.Relationships.List.Add oRelationship
            Next
            For lIndex = 0 To 3
                oCurvePositions(lIndex).RenderName
            Next
            DiagramRef.FileIOs.WriteFile
                        
        Case vbKeyRight
            lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            For lIndex = 0 To 3
                Set oCurvePositions(lIndex) = New Position
                
                With oCurvePositions(lIndex)
                    Set .DiagramRef = DiagramRef
                    Set .ParserRef = oParsePosition
                    DiagramRef.Positions.List.Add oCurvePositions(lIndex)
                End With
            Next
            oCurvePositions(0).Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
            oCurvePositions(0).Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
            
            oCurvePositions(1).Pos.X = oCurvePositions(0).Pos.X + 20
            oCurvePositions(1).Pos.Y = oCurvePositions(0).Pos.Y
            
            oCurvePositions(2).Pos.X = oCurvePositions(1).Pos.X + 20
            oCurvePositions(2).Pos.Y = oCurvePositions(1).Pos.Y - 20
            
            oCurvePositions(3).Pos.X = oCurvePositions(2).Pos.X
            oCurvePositions(3).Pos.Y = oCurvePositions(2).Pos.Y - 20

            For lIndex = 0 To 2
                Set oRelationship = New Relationship
                Set oRelationship.DiagramRef = DiagramRef
                Set oRelationship.FromPos = oCurvePositions(lIndex)
                Set oRelationship.ToPos = oCurvePositions(lIndex + 1)
                If lIndex = 1 Then
                    oRelationship.Style = 1
                End If
                oRelationship.ColourIndeces = Array(mlRecentColourIndex)
                oRelationship.RenderRelationship
                DiagramRef.Relationships.List.Add oRelationship
                DiagramRef.Positions.RenderAll
            Next
            For lIndex = 0 To 3
                oCurvePositions(lIndex).RenderName
            Next
            DiagramRef.FileIOs.WriteFile
            
        Case vbKeyUp
           lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            For lIndex = 0 To 3
                Set oCurvePositions(lIndex) = New Position
                
                With oCurvePositions(lIndex)
                    Set .DiagramRef = DiagramRef
                    Set .ParserRef = oParsePosition
                    DiagramRef.Positions.List.Add oCurvePositions(lIndex)
                End With
            Next
            oCurvePositions(0).Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
            oCurvePositions(0).Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
            
            oCurvePositions(1).Pos.X = oCurvePositions(0).Pos.X - 20
            oCurvePositions(1).Pos.Y = oCurvePositions(0).Pos.Y
            
            oCurvePositions(2).Pos.X = oCurvePositions(1).Pos.X - 20
            oCurvePositions(2).Pos.Y = oCurvePositions(1).Pos.Y - 20
            
            oCurvePositions(3).Pos.X = oCurvePositions(2).Pos.X
            oCurvePositions(3).Pos.Y = oCurvePositions(2).Pos.Y - 20

            For lIndex = 0 To 2
                Set oRelationship = New Relationship
                Set oRelationship.DiagramRef = DiagramRef
                Set oRelationship.FromPos = oCurvePositions(lIndex)
                Set oRelationship.ToPos = oCurvePositions(lIndex + 1)
                If lIndex = 1 Then
                    oRelationship.Style = 1
                End If
                oRelationship.ColourIndeces = Array(mlRecentColourIndex)
                oRelationship.RenderRelationship
                DiagramRef.Relationships.List.Add oRelationship
                DiagramRef.Positions.RenderAll
            Next
            For lIndex = 0 To 3
                oCurvePositions(lIndex).RenderName
            Next
            DiagramRef.FileIOs.WriteFile

        Case vbKeyDown
           lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            For lIndex = 0 To 3
                Set oCurvePositions(lIndex) = New Position
                
                With oCurvePositions(lIndex)
                    Set .DiagramRef = DiagramRef
                    Set .ParserRef = oParsePosition
                    DiagramRef.Positions.List.Add oCurvePositions(lIndex)
                End With
            Next
            oCurvePositions(0).Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
            oCurvePositions(0).Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
            
            oCurvePositions(1).Pos.X = oCurvePositions(0).Pos.X + 20
            oCurvePositions(1).Pos.Y = oCurvePositions(0).Pos.Y
            
            oCurvePositions(2).Pos.X = oCurvePositions(1).Pos.X + 20
            oCurvePositions(2).Pos.Y = oCurvePositions(1).Pos.Y + 20
            
            oCurvePositions(3).Pos.X = oCurvePositions(2).Pos.X
            oCurvePositions(3).Pos.Y = oCurvePositions(2).Pos.Y + 20

            For lIndex = 0 To 2
                Set oRelationship = New Relationship
                Set oRelationship.DiagramRef = DiagramRef
                Set oRelationship.FromPos = oCurvePositions(lIndex)
                Set oRelationship.ToPos = oCurvePositions(lIndex + 1)
                If lIndex = 1 Then
                    oRelationship.Style = 1
                End If
                oRelationship.ColourIndeces = Array(mlRecentColourIndex)
                oRelationship.RenderRelationship
                DiagramRef.Relationships.List.Add oRelationship
                DiagramRef.Positions.RenderAll
            Next
            For lIndex = 0 To 3
                oCurvePositions(lIndex).RenderName
            Next
            DiagramRef.FileIOs.WriteFile
        Case vbKeyReturn
            DiagramRef.RenderAndSave
    End Select
End Sub

Private Sub Form_Load()
    BackColour = Me.BackColor
    
    DiagramRef.RenderAndSave
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    
    If Button = vbLeftButton Then
        mnPreviousX = X
        mnPreviousY = Y
        Set oSelectedPosition = DiagramRef.Positions.FindPosition((X - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (Y - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
        If Not oSelectedPosition Is Nothing Then
            oSelectedPosOffsetX = (oSelectedPosition.Pos.X * DiagramRef.Zoom + DiagramRef.TopLeft.X) - X
            oSelectedPosOffsetY = (oSelectedPosition.Pos.Y * DiagramRef.Zoom + DiagramRef.TopLeft.Y) - Y
            If Not oSelectedPositions Is Nothing Then
                bDragGroupSelected = False
                For Each oPosition In oSelectedPositions
                    If oPosition Is oSelectedPosition Then
                        bDragGroupSelected = True
                    End If
                Next
            End If
        Else
            bDragGroupSelected = False
            If Shift = 0 Then
                Set oSelectedPositions = Nothing
                Set oInitialLink = Nothing
                Set oFinalLink = Nothing
                DiagramRef.Positions.RenderAll
            Else
                nBoxStartX = X
                nBoxStartY = Y
            End If
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPositionA As Position
    Dim nSelectedX As Single
    Dim nSelectedY As Single
    Dim oTempPosition As Position
    
    mnMouseX = X
    mnMouseY = Y
    If Button = vbLeftButton Then
        If Not oSelectedPosition Is Nothing Then
            If bDragGroupSelected Then
                If Not oSelectedPositions Is Nothing Then
                    nSelectedX = oSelectedPosition.Pos.X
                    nSelectedY = oSelectedPosition.Pos.Y
                    For Each oPositionA In oSelectedPositions
                        oPositionA.ClearName
                        oPositionA.Pos.X = oPositionA.Pos.X + ((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + oSelectedPosOffsetX - nSelectedX)
                        oPositionA.Pos.Y = oPositionA.Pos.Y + ((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + oSelectedPosOffsetY - nSelectedY)
                        oPositionA.RenderName &HFFC0C0
                    Next
                End If
            Else
                If Shift = 3 Then
                    Set oSelectedPosition = oSelectedPosition.Copy
                End If
                
                oSelectedPosition.ClearName
                oSelectedPosition.Pos.X = (mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + oSelectedPosOffsetX
                oSelectedPosition.Pos.Y = (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + oSelectedPosOffsetY
                oSelectedPosition.RenderName &HFFC0C0
            End If
        Else
            If Shift = 0 Then
                If mnPreviousX <> 0 Or mnPreviousY <> 0 Then
                    DiagramRef.TopLeft.X = DiagramRef.TopLeft.X - (mnPreviousX - X)
                    DiagramRef.TopLeft.Y = DiagramRef.TopLeft.Y - (mnPreviousY - Y)

                    DiagramRef.RenderAndSave
                    mnPreviousX = X
                    mnPreviousY = Y
                End If
            Else
                If Not oSelectedPositions Is Nothing Then
                    For Each oPositionA In oSelectedPositions
                        oPositionA.RenderName
                    Next
                End If
                Set oSelectedPositions = DiagramRef.Positions.FindPositions((nBoxStartX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (nBoxStartY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom, (X - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (Y - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                For Each oPositionA In oSelectedPositions
                    oPositionA.RenderName &HFFC0C0
                Next
                
                Me.ForeColor = vbWhite
                Me.DrawMode = vbXorPen
                Me.DrawStyle = vbDash
                Me.FillStyle = 1
                If mnPreviousX <> 0 Or mnPreviousY <> 0 Then
                    Me.Line (nBoxStartX, nBoxStartY)-(mnPreviousX, mnPreviousY), , B
                End If
                Me.Line (nBoxStartX, nBoxStartY)-(X, Y), , B

                Me.DrawMode = vbCopyPen
                Me.DrawStyle = vbSolid
                
                mnPreviousX = X
                mnPreviousY = Y
            End If
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position

    If Shift = 1 Then
        Me.Line (nBoxStartX, nBoxStartY)-(mnPreviousX, mnPreviousY), , B
    End If
    
    If Not oSelectedPosition Is Nothing Then
        If Shift <> 1 And Shift <> 3 Then
            oSelectedPosition.Pos.X = Int((oSelectedPosition.Pos.X + 10) / 20) * 20
            oSelectedPosition.Pos.Y = Int((oSelectedPosition.Pos.Y + 10) / 20) * 20
                    
            If Not oSelectedPositions Is Nothing Then
                For Each oPosition In oSelectedPositions
                    oPosition.Pos.X = Int((oPosition.Pos.X + 10) / 20) * 20
                    oPosition.Pos.Y = Int((oPosition.Pos.Y + 10) / 20) * 20
                Next
            End If
        End If
        
        If Shift = 3 Then
            Set oPosition = DiagramRef.Positions.FindPosition((X - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (Y - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                oPosition.Name = oSelectedPosition.Name
                oPosition.RenderName &HFFC0C0
            End If
            Set oSelectedPosition = Nothing
        End If
        
        If Shift = 2 Then
            If oSelectedPositions Is Nothing Then
                Set oSelectedPositions = New Collection
            End If
            
            oSelectedPositions.Add oSelectedPosition
            oSelectedPosition.RenderName &HFFC0C0
        Else
            Set oSelectedPosition = Nothing
        End If
        
        DiagramRef.RenderAndSave

        If Not oSelectedPositions Is Nothing Then
            For Each oPosition In oSelectedPositions
                oPosition.RenderName &HFFC0C0
            Next
        End If
    End If
    
    mnPreviousX = 0
    mnPreviousY = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DiagramRef.FileIOs.WriteFile
End Sub

Private Sub scrColourComponent_Change(Index As Integer)
    Dim yColour(3) As Byte
    Dim vColours As Variant
    Dim lColour As Long
    
    yColour(0) = scrColourComponent(0).Value
    yColour(1) = scrColourComponent(1).Value
    yColour(2) = scrColourComponent(2).Value
    
    CopyMemory lColour, yColour(0), &H4&

    vColours = DiagramRef.Colours
    vColours(mlRecentColourIndex) = lColour
    txtRGB.Text = "#" & Pad(Hex$(yColour(0))) & Pad(Hex$(yColour(1))) & Pad(Hex$(yColour(2)))
    DiagramRef.Colours = vColours
    DiagramRef.RenderAndSave
End Sub

Private Function Pad(sString As String) As String
    Pad = String$(2 - Len(sString), "0") & sString
End Function

Public Sub Watermark()
    Dim nX As Single
    Dim nY As Single
    
    Me.ForeColor = DiagramRef.Colours(17)
    For nX = 0 To 1600 Step 160
        For nY = 0 To 1200 Step 15
            CurrentX = nX + ((nY \ 15) Mod 2) * 80
            CurrentY = nY
            
            Print "Copyright " & Chr$(169) & " Guillermo Phillips 2007"
            Print
        Next
    Next
    
End Sub

Private Sub txtRGB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        scrColourComponent(0).Value = HexToDecimal(Mid$(txtRGB.Text, 2, 2))
        scrColourComponent(1).Value = HexToDecimal(Mid$(txtRGB.Text, 4, 2))
        scrColourComponent(2).Value = HexToDecimal(Mid$(txtRGB.Text, 6, 2))
    End If
End Sub

Private Function HexToDecimal(sHex As String) As Long
    Const sFigures = "0123456789ABCDEF"
    Dim lIndex  As Long
    Dim lMultiplier As Long
    
    lMultiplier = 1
    For lIndex = Len(sHex) To 1 Step -1
        HexToDecimal = HexToDecimal + (InStr(sFigures, Mid$(sHex, lIndex, 1)) - 1) * lMultiplier
        lMultiplier = lMultiplier * 16
    Next
End Function

