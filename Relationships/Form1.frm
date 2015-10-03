VERSION 5.00
Begin VB.Form Canvas 
   AutoRedraw      =   -1  'True
   Caption         =   "Relationships"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Canvas"
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

Private oSelectedPosition As Position
Private oSelectedPosOffsetX As Single
Private oSelectedPosOffsetY As Single

Private oInitialLink As Position
Private oFinalLink As Position

Public BackColour As Long


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
            Set oPosition.CanvasRef = Me
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
        Case vbKeyL
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
                        Set oRelationship.CanvasRef = Me
                        Set oRelationship.FromPos = oInitialLink
                        Set oRelationship.ToPos = oFinalLink
                        Set oRelationship.PositionListRef = Positions.List
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
    DoDefinition
    Initialise Me
    BackColour = Me.BackColor
    Relations.RenderAll
    Positions.RenderAll
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static oPosition As Position
    If Button = vbLeftButton Then
        Set oSelectedPosition = Positions.FindPosition(X - sTopLeftX, Y - sTopLeftY)
        If Not oSelectedPosition Is Nothing Then
            oSelectedPosOffsetX = oSelectedPosition.PosX - X
            oSelectedPosOffsetY = oSelectedPosition.PosY - Y
        Else
            Set oInitialLink = Nothing
            Set oFinalLink = Nothing
            Positions.RenderAll
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PreviousX As Single
    Static PreviousY As Single
    
    sMouseX = X
    sMouseY = Y
    If Button = vbLeftButton Then
        If Not oSelectedPosition Is Nothing Then
            oSelectedPosition.ClearName
            oSelectedPosition.PosX = sMouseX + oSelectedPosOffsetX
            oSelectedPosition.PosY = sMouseY + oSelectedPosOffsetY
            oSelectedPosition.RenderName
        Else
            sTopLeftX = sTopLeftX - (PreviousX - X)
            sTopLeftY = sTopLeftY - (PreviousY - Y)
            Cls
            Relations.RenderAll
            Positions.RenderAll
        End If
    End If
    PreviousX = X
    PreviousY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not oSelectedPosition Is Nothing Then
        Set oSelectedPosition = Nothing
        Cls
        Relations.RenderAll
        Positions.RenderAll
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeInitialise
End Sub
