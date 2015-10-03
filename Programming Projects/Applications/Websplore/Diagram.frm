VERSION 5.00
Begin VB.Form Diagram 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Files"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   11865
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
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

Private nMouseX As Single
Private nMouseY As Single

Public nTopLeftX As Single
Public nTopLeftY As Single

Public nZoom As Single

Private oSelectedPosition As Position
Private nSelectedPosOffsetX As Single
Private nSelectedPosOffsetY As Single

Private nBoxStartX As Single
Private nBoxStartY As Single

Public BackColour As Long

Public Positions As PositionList
Public Relations As RelationshipList
Public FileIOs As FileIO

Private bDragGroupSelected As Boolean

Private vColours As Variant

Private Sub Form_DblClick()
    Dim oPosition As Position
    
    Set oPosition = Positions.FindPosition(nMouseX - nTopLeftX, nMouseY - nTopLeftY)
    
    If Not oPosition Is Nothing Then
        ShellExecute Me.hdc, "Open", oPosition.Path, "", "", SW_SHOWNORMAL
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oPosition As Position
    Dim oRelationship As Relationship
    
    Set oPosition = Positions.FindPosition(nMouseX - nTopLeftX, nMouseY - nTopLeftY)
    
    Select Case KeyCode
        Case vbKeyAdd
            For Each oPosition In Positions.SelectedList
                oPosition.Radius = oPosition.Radius + 1
                'Positions.Highlight oPosition, True
            Next
        Case vbKeySubtract
            If Not oPosition Is Nothing Then
                oPosition.Radius = oPosition.Radius - 1
                If oPosition.Radius < 5 Then
                    oPosition.Radius = 5
                End If
                Positions.Highlight oPosition, True
            End If
        Case vbKeyR
            If Not oPosition Is Nothing Then
                oPosition.Orientation = (oPosition.Orientation + 1) Mod 8
                Positions.Highlight oPosition, True
                Redraw
            End If
        Case vbKeyDelete
            If Positions.SelectedList.Count > 0 Then
                For Each oPosition In Positions.SelectedList
                    Relations.RemoveRelationshipWithReference oPosition
                    Positions.RemovePosition oPosition
                    FileIOs.WriteFile
                Next
                Redraw
            End If
        Case vbKeyL
            If Positions.SelectedList.Count = 2 Then
                Set oRelationship = Relations.FindRelationship(Positions.SelectedList(1), Positions.SelectedList(2))
                If Not oRelationship Is Nothing Then
                    oRelationship.AngleIndex = (oRelationship.AngleIndex + 1) Mod 7
                    Redraw
                End If
            End If
        Case vbKey0 To vbKey9
            If oPosition Is Nothing Then
                NewString.txtString = ""
                NewString.Show vbModal
                Set oPosition = New Position
                oPosition.Name = NewString.txtString
                oPosition.PosX = nMouseX - nTopLeftX
                oPosition.PosY = nMouseY - nTopLeftY
                oPosition.Snap
                oPosition.Colour = vColours(Shift * 10 + KeyCode - 48)
                Set oPosition.DiagramRef = Me
                Set oPosition.ParserRef = oParsePosition
                Positions.List.Add oPosition
                oPosition.RenderName
                Positions.Highlight oPosition, True
                FileIOs.WriteFile
            Else
                Positions.Highlight oPosition, True
                If Positions.SelectedList.Count = 2 Then
                    Set oRelationship = Relations.FindRelationship(Positions.SelectedList(1), Positions.SelectedList(2))
                    If oRelationship Is Nothing Then
                        Set oRelationship = New Relationship
                        Set oRelationship.DiagramRef = Me
                        Set oRelationship.FromPos = Positions.SelectedList(1)
                        Set oRelationship.ToPos = Positions.SelectedList(2)
                        Set oRelationship.PositionListRef = Positions.List
                        oRelationship.Colour = vColours(Shift * 10 + KeyCode - 48)
                        oRelationship.RenderRelationship
                        Relations.List.Add oRelationship
                        Positions.Highlight Positions.SelectedList(1), False
                        Positions.Highlight Positions.SelectedList(1), False
                        Positions.RenderAll
                    Else
                        Relations.RemoveRelationship oRelationship
                        Positions.Highlight Positions.SelectedList(1), False
                        Positions.Highlight Positions.SelectedList(1), False
                        Redraw
                    End If
                End If
            End If
        Case vbKeyE
            If Positions.SelectedList.Count = 1 Then
                Set oPosition = Positions.SelectedList(1)
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
    vColours = Array(RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 255, 0), RGB(255, 255, 0), RGB(0, 0, 255), RGB(255, 0, 255), RGB(0, 255, 255), RGB(255, 255, 255), RGB(128, 128, 128), &H80FF&, &H404080, &H40C0&, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    
    nMouseX = X
    nMouseY = Y
    If Button = vbLeftButton Then
        Set oSelectedPosition = Positions.FindPosition(X - nTopLeftX, Y - nTopLeftY)
        If Not oSelectedPosition Is Nothing Then
            nSelectedPosOffsetX = oSelectedPosition.PosX - X
            nSelectedPosOffsetY = oSelectedPosition.PosY - Y
            If Positions.SelectedList.Count > 0 Then
                bDragGroupSelected = False
                For Each oPosition In Positions.SelectedList
                    If oPosition Is oSelectedPosition Then
                        bDragGroupSelected = True
                    End If
                Next
            Else
                If Shift = 2 Then
'                    For Each oPosition In Positions.List
'                        If oPosition.Highlighted Then
'                            bDragGroupSelected = True
'                            If oSelectedPositions Is Nothing Then
'                                Set oSelectedPositions = New Collection
'                            End If
'                            oSelectedPositions.Add oPosition
'                        End If
'                    Next
                Else
                    Positions.RemoveHighlights
                End If
                Positions.Highlight oSelectedPosition, True
            End If
        Else
            Positions.RemoveHighlights
            bDragGroupSelected = False
            If Shift = 0 Then
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
    Dim oPosition As Position
    Dim nSelectedX As Single
    Dim nSelectedY As Single
    
    nMouseX = X
    nMouseY = Y
    If Button = vbLeftButton Then
        If Not oSelectedPosition Is Nothing Then
            If bDragGroupSelected Then
                If Positions.SelectedList.Count > 0 Then
                    nSelectedX = oSelectedPosition.PosX
                    nSelectedY = oSelectedPosition.PosY
                    For Each oPosition In Positions.SelectedList
                        oPosition.ClearName
                        oPosition.PosX = oPosition.PosX + (nMouseX + nSelectedPosOffsetX - nSelectedX)
                        oPosition.PosY = oPosition.PosY + (nMouseY + nSelectedPosOffsetY - nSelectedY)
                        oPosition.RenderName RGB(255 * Rnd, 255 * Rnd, 255 * Rnd) '&HFFC0C0
                    Next
                End If
            Else
                oSelectedPosition.ClearName
                oSelectedPosition.PosX = nMouseX + nSelectedPosOffsetX
                oSelectedPosition.PosY = nMouseY + nSelectedPosOffsetY
                oSelectedPosition.RenderName &HFFC0C0
            End If
        Else
            If Shift = 0 Then
                If PreviousX <> 0 Or PreviousY <> 0 Then
                    nTopLeftX = nTopLeftX - (PreviousX - X)
                    nTopLeftY = nTopLeftY - (PreviousY - Y)
                    Redraw
                End If
            Else
                Me.ForeColor = vbWhite
                Me.DrawMode = vbXorPen
                Me.DrawStyle = vbDash
                Me.FillStyle = 1
                Me.Line (nBoxStartX, nBoxStartY)-(PreviousX, PreviousY), , B
                Me.Line (nBoxStartX, nBoxStartY)-(X, Y), , B

                Me.DrawMode = vbCopyPen
                Me.DrawStyle = vbSolid
                
                For Each oPosition In Positions.SelectedList
                    oPosition.Highlight False
                Next
                
                Set Positions.SelectedList = Positions.FindPositions(nBoxStartX - nTopLeftX, nBoxStartY - nTopLeftY, X - nTopLeftX, Y - nTopLeftY)
                For Each oPosition In Positions.SelectedList
                    oPosition.Highlight True
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
        If Shift <> 1 Then
            oSelectedPosition.Snap
        End If
        Set oSelectedPosition = Nothing

        For Each oPosition In Positions.SelectedList
            If Shift <> 1 Then
                oPosition.Snap
            End If
        Next
        For Each oPosition In Positions.SelectedList
            oPosition.Highlight True
        Next
    End If
    Redraw
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


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    Dim oFSO As New FileSystemObject
    Dim oFile As File
    Dim oFolder As Folder
    Dim oPic As IPictureDisp
    Dim vFile As Variant
    Dim lIndex As Long
    
    On Error Resume Next
    
    For Each vFile In Data.Files
        Set oPosition = New Position
        
        If oFSO.FileExists(vFile) Then
            Set oFile = oFSO.GetFile(vFile)
            oPosition.Name = oFile.Name
            oPosition.Path = vFile
        ElseIf oFSO.FolderExists(vFile) Then
            Set oFolder = oFSO.GetFolder(vFile)
            oPosition.Name = oFolder.Name
            oPosition.Path = vFile
        Else
            oPosition.Name = vFile
        End If
        
        oPosition.PosX = nMouseX - nTopLeftX
        oPosition.PosY = (nMouseY - nTopLeftY) + 20 * lIndex
        oPosition.Snap
        oPosition.Colour = vColours(0)
        Set oPosition.DiagramRef = Me
        Set oPosition.ParserRef = oParsePosition
        Positions.List.Add oPosition
        oPosition.RenderName
        FileIOs.WriteFile
        lIndex = lIndex + 1
    Next
End Sub

Private Sub Redraw()
    Cls
    Relations.RenderAll
    Positions.RenderAll
    FileIOs.WriteFile
End Sub
