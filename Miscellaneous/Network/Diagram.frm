VERSION 5.00
Begin VB.Form Diagram 
   AutoRedraw      =   -1  'True
   Caption         =   "Relationships"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   11865
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Diagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TopLeft As Single
Private TopRight As Single

Private sMouseX As Single
Private sMouseY As Single

Public sTopLeftX As Single
Public sTopLeftY As Single

Private oSelectedPositions As Collection

Private oSelectedPosition As Position
Private oSelectedPosOffsetX As Single
Private oSelectedPosOffsetY As Single

Private nBoxStartX As Single
Private nBoxStartY As Single

Private oInitialLink As Position
Private oFinalLink As Position

Public BackColour As Long

Public Positions As PositionList
Public Relations As RelationshipList
Public FileIOs As FileIO

Private bDragGroupSelected As Boolean

Private vColours As Variant

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oPosition As Position
    Dim oRelationship As Relationship
    
    Select Case KeyCode
        Case vbKeyN
            NewString.txtString = ""
            NewString.Show vbModal
            Set oPosition = New Position
            oPosition.Name = NewString.txtString
            oPosition.PosX = sMouseX - sTopLeftX
            oPosition.PosY = sMouseY - sTopLeftY
            Set oPosition.DiagramRef = Me
            Set oPosition.ParserRef = oParsePosition
            Positions.List.Add oPosition
            oPosition.RenderName
            FileIOs.WriteFile
        Case vbKeyDelete
            Set oPosition = Positions.FindPosition(sMouseX - sTopLeftX, sMouseY - sTopLeftY)
            If Not oPosition Is Nothing Then
                Relations.RemoveRelationshipWithReference oPosition
            End If
            Positions.RemovePosition oPosition
            Cls
            Relations.RenderAll
            Positions.RenderAll
            FileIOs.WriteFile
        Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKey0
            If oInitialLink Is Nothing Then
                Set oInitialLink = Positions.FindPosition(sMouseX - sTopLeftX, sMouseY - sTopLeftY)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName vbGreen
                End If
            Else
                Set oFinalLink = Positions.FindPosition(sMouseX - sTopLeftX, sMouseY - sTopLeftY)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName vbGreen
                    
                    Set oRelationship = Relations.FindRelationship(oInitialLink, oFinalLink)
                    If oRelationship Is Nothing Then
                        Set oRelationship = New Relationship
                        Set oRelationship.DiagramRef = Me
                        Set oRelationship.FromPos = oInitialLink
                        Set oRelationship.ToPos = oFinalLink
                        Set oRelationship.PositionListRef = Positions.List
                        oRelationship.Colour = vColours(KeyCode - 48)
                        oRelationship.RenderRelationship
                        Relations.List.Add oRelationship
                        Positions.RenderAll
                    Else
                        Relations.RemoveRelationship oRelationship
                        Cls
                        Relations.RenderAll
                        Positions.RenderAll
                    End If
                    Set oInitialLink = Nothing
                    Set oFinalLink = Nothing
                    FileIOs.WriteFile
                Else
                    oInitialLink.RenderName
                    Set oInitialLink = Nothing
                End If
            End If
        
        Case vbKeyE
            Set oPosition = Positions.FindPosition(sMouseX - sTopLeftX, sMouseY - sTopLeftY)
            If Not oPosition Is Nothing Then
                oPosition.ClearName
                NewString.txtString = oPosition.Name
                NewString.Show vbModal
                oPosition.Name = NewString.txtString
                oPosition.RenderName
                FileIOs.WriteFile
            End If
    End Select
End Sub

Private Sub Form_Load()
    BackColour = Me.BackColor
    Relations.RenderAll
    Positions.RenderAll
    vColours = Array(RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 255, 0), RGB(255, 255, 0), RGB(0, 0, 255), RGB(255, 0, 255), RGB(0, 255, 255), RGB(255, 255, 255), RGB(128, 128, 128), RGB(0, 0, 0))
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    
    If Button = vbLeftButton Then
        Set oSelectedPosition = Positions.FindPosition(X - sTopLeftX, Y - sTopLeftY)
        If Not oSelectedPosition Is Nothing Then
            oSelectedPosOffsetX = oSelectedPosition.PosX - X
            oSelectedPosOffsetY = oSelectedPosition.PosY - Y
            If Not oSelectedPositions Is Nothing Then
                bDragGroupSelected = False
                For Each oPosition In oSelectedPositions
                    If oPosition Is oSelectedPosition Then
                        bDragGroupSelected = True
                    End If
                Next
            End If
        Else
            If Shift = 0 Then
                Set oSelectedPositions = Nothing
                Set oInitialLink = Nothing
                Set oFinalLink = Nothing
                Positions.RenderAll
            Else
                nBoxStartX = X
                nBoxStartY = Y
            End If
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PreviousX As Single
    Static PreviousY As Single
    Dim oPositionA As Position
    
    sMouseX = X
    sMouseY = Y
    If Button = vbLeftButton Then
        If Not oSelectedPosition Is Nothing Then
            If bDragGroupSelected Then
                If Not oSelectedPositions Is Nothing Then
                    For Each oPositionA In oSelectedPositions
                        oPositionA.ClearName
                        oPositionA.PosX = oPositionA.PosX + (sMouseX + oSelectedPosOffsetX - oSelectedPosition.PosX)
                        oPositionA.PosY = oPositionA.PosY + (sMouseY + oSelectedPosOffsetY - oSelectedPosition.PosY)
                        oPositionA.RenderName vbGreen
                    Next
                End If
            Else
                oSelectedPosition.ClearName
                oSelectedPosition.PosX = sMouseX + oSelectedPosOffsetX
                oSelectedPosition.PosY = sMouseY + oSelectedPosOffsetY
                oSelectedPosition.RenderName vbGreen
            End If
        Else
            If Shift = 0 Then
                If PreviousX <> 0 Or PreviousY <> 0 Then
                    sTopLeftX = sTopLeftX - (PreviousX - X)
                    sTopLeftY = sTopLeftY - (PreviousY - Y)
                    Cls
                    Relations.RenderAll
                    Positions.RenderAll
                End If
            Else
                Me.DrawMode = vbXorPen
                Me.DrawStyle = vbDash
                Me.Line (nBoxStartX, nBoxStartY)-(PreviousX, nBoxStartY)
                Me.Line (PreviousX, nBoxStartY)-(PreviousX, PreviousY)
                Me.Line (PreviousX, PreviousY)-(nBoxStartX, PreviousY)
                Me.Line (nBoxStartX, PreviousY)-(nBoxStartX, nBoxStartY)
                
                Me.Line (nBoxStartX, nBoxStartY)-(X, nBoxStartY)
                Me.Line (X, nBoxStartY)-(X, Y)
                Me.Line (X, Y)-(nBoxStartX, Y)
                Me.Line (nBoxStartX, Y)-(nBoxStartX, Y)
                                
                Me.DrawMode = vbCopyPen
                Me.DrawStyle = vbSolid
                
                Set oSelectedPositions = Positions.FindPositions(nBoxStartX - sTopLeftX, nBoxStartY - sTopLeftY, X - sTopLeftX, Y - sTopLeftY)
                For Each oPositionA In oSelectedPositions
                    oPositionA.RenderName vbGreen
                Next
            End If
        End If
    End If
    PreviousX = X
    PreviousY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    If Not oSelectedPosition Is Nothing Then
        Set oSelectedPosition = Nothing
        Cls
        Relations.RenderAll
        Positions.RenderAll
        If Not oSelectedPositions Is Nothing Then
            For Each oPosition In oSelectedPositions
                oPosition.RenderName vbGreen
            Next
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FileIOs.WriteFile
End Sub


Public Sub Initialise(sFilePath As String)
    Set Positions = New PositionList
    Set Relations = New RelationshipList
    
    Set FileIOs = New FileIO
    Set FileIOs.DiagramRef = Me
    Set FileIOs.PositionsRef = Positions
    Set FileIOs.RelationsRef = Relations
    If sFilePath <> "" Then
        FileIOs.FileStore = sFilePath
    Else
        FileIOs.FileStore = App.Path & "\diagram.txt"
    End If
    FileIOs.ReadFile
End Sub


