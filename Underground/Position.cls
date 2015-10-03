VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
    Data(15) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long

Public Reference As String
Public Name As String
Public ColourIndex As Long
Public Orientation As Long
Public Shape As Long
Public Proximity As Long

Public Pos As New Vector

Public DiagramRef As Diagram
Public ParserRef As IParseObject

Private Const Margin As Single = 1

Public Function Copy() As Position
    Set Copy = New Position
    
    With Copy
        .Name = Me.Name
        .ColourIndex = Me.ColourIndex
        .Orientation = Me.Orientation
        .Shape = Me.Shape
        Set .Pos = New Vector
        .Pos.Copy Me.Pos
        Set .DiagramRef = Me.DiagramRef
        Set .ParserRef = Me.ParserRef
    End With
End Function

Public Function LeftTextPosition()
    LeftTextPosition = Pos.X - DiagramRef.PaperRef.TextWidth(Name) / 2
End Function

Public Function RightTextPosition()
    RightTextPosition = DiagramRef.PaperRef.TextWidth(Name) / 2 + Pos.X
End Function

Public Function TopTextPosition()
    TopTextPosition = Pos.Y - DiagramRef.PaperRef.TextHeight(Name) / 2
End Function

Public Function BottomTextPosition()
    BottomTextPosition = Pos.Y + DiagramRef.PaperRef.TextHeight(Name) / 2
End Function

Public Function Height()
    Height = DiagramRef.PaperRef.TextHeight(Name)
End Function

Public Function Radius()
    Radius = Margin + Height / 2 - 1
End Function

Public Function RenderName(Optional iColour As Long = -1)
    Dim nTextX As Single
    Dim nTextY As Single
    Dim lFillColour As Long
    Dim bUnfilled As Boolean
    
    With DiagramRef.PaperRef
        .Font.Size = FontSize * DiagramRef.Zoom
        If iColour = -1 Then
            iColour = DiagramRef.Colours(ColourIndex)
        End If
        If .mbShowCircles Then
            lFillColour = &HFFC0C0
        Else
            lFillColour = iColour
        End If
        
        .FillStyle = 0
        If Name <> "" Then
            .DrawWidth = 1
            If lFillColour <> 0 Then
                .FillColor = .BackColor
                .ForeColor = .BackColor
            Else
                .FillColor = vbWhite
                .ForeColor = vbWhite
            End If
            
            bUnfilled = (Shape And 1) = 1
            Select Case Shape
                Case 0, 1 ' Circle
                    DiagramRef.PaperRef.Circle (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom), OuterCircleRadius * DiagramRef.Zoom
                    .ForeColor = lFillColour
                    If bUnfilled Then
                        .FillColor = .BackColor
                    Else
                        .FillColor = lFillColour
                    End If
                    DiagramRef.PaperRef.Circle (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom), InnerCircleRadius * DiagramRef.Zoom
                    .DrawWidth = 1
                Case 2, 3 ' Square
                    DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (OuterCircleRadius) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - (OuterCircleRadius) * DiagramRef.Zoom)-Step((OuterCircleRadius) * 2 * DiagramRef.Zoom, (OuterCircleRadius) * 2 * DiagramRef.Zoom), lFillColour, BF
                    If bUnfilled Then
                        DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (InnerCircleRadius) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - (InnerCircleRadius) * DiagramRef.Zoom)-Step((InnerCircleRadius) * 2 * DiagramRef.Zoom, (InnerCircleRadius) * 2 * DiagramRef.Zoom), .BackColor, BF
                    End If
                Case 4, 5 ' Vertical Rectangle
                    DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (OuterCircleRadius - 2) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + (Pos.Y - 1) * DiagramRef.Zoom - (OuterCircleRadius - 2) * DiagramRef.Zoom)-Step((OuterCircleRadius - 2) * 2 * DiagramRef.Zoom, ((OuterCircleRadius) * 2 - 2) * DiagramRef.Zoom), lFillColour, BF
                    If bUnfilled Then
                        DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (InnerCircleRadius - 2) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + (Pos.Y - 1) * DiagramRef.Zoom - (InnerCircleRadius - 2) * DiagramRef.Zoom)-Step((InnerCircleRadius - 2) * 2 * DiagramRef.Zoom, ((InnerCircleRadius) * 2 - 2) * DiagramRef.Zoom), .BackColor, BF
                    End If
                Case 6, 7 ' Horizontal Rectangle
                    DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (OuterCircleRadius) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - (OuterCircleRadius - 2) * DiagramRef.Zoom)-Step(((OuterCircleRadius) * 2 - 1) * DiagramRef.Zoom, (OuterCircleRadius - 2) * 2 * DiagramRef.Zoom), lFillColour, BF
                    If bUnfilled Then
                        DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - (InnerCircleRadius) * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - (InnerCircleRadius - 2) * DiagramRef.Zoom)-Step(((InnerCircleRadius) * 2 - 1) * DiagramRef.Zoom, (InnerCircleRadius - 2) * 2 * DiagramRef.Zoom), .BackColor, BF
                    End If
            End Select
        Else
            If .mbShowCircles Then
                .DrawWidth = 1
                .ForeColor = lFillColour
                .FillColor = lFillColour
                DiagramRef.PaperRef.Circle (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom), LineWidth / 2 * DiagramRef.Zoom - 1
                .DrawWidth = 1
            End If
        End If
        
        If Trim$(Name) <> "" Then
            Select Case Orientation
                Case 0, 4
                    nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - Height / 2
                Case 1, 2, 3
                    nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom + TextSeparation * Proximity * DiagramRef.Zoom
                Case 5, 6, 7
                    nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - Height - TextSeparation * Proximity * DiagramRef.Zoom
            End Select
            
            Select Case Orientation
                Case 7, 0, 1
                    nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom + TextSeparation * Proximity * DiagramRef.Zoom
                Case 2, 6
                    nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - .TextWidth(Name) / 2
                Case 3, 4, 5
                    nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - .TextWidth(Name) - TextSeparation * Proximity * DiagramRef.Zoom
            End Select

            'DiagramRef.PaperRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), .BackColor, BF
            .CurrentX = nTextX
            .CurrentY = nTextY
            
            .ForeColor = DiagramRef.Colours(19)
            DiagramRef.PaperRef.Print Name
        End If
    End With
End Function

Public Function ClearName()

    Dim nTextX As Single
    Dim nTextY As Single
    
    With DiagramRef.PaperRef
        .FillColor = .BackColor
        .ForeColor = .BackColor
        DiagramRef.PaperRef.Circle (DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom, DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom), OuterCircleRadius * DiagramRef.Zoom
        
        Select Case Orientation
            Case 0, 4
                nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - Height / 2
            Case 1, 2, 3
                nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom + TextSeparation * Proximity * DiagramRef.Zoom
            Case 5, 6, 7
                nTextY = DiagramRef.TopLeft.Y + Pos.Y * DiagramRef.Zoom - Height - TextSeparation * Proximity * DiagramRef.Zoom
        End Select
        
        Select Case Orientation
            Case 7, 0, 1
                nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom + TextSeparation * Proximity * DiagramRef.Zoom
            Case 2, 6
                nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - .TextWidth(Name) / 2
            Case 3, 4, 5
                nTextX = DiagramRef.TopLeft.X + Pos.X * DiagramRef.Zoom - .TextWidth(Name) - TextSeparation * Proximity * DiagramRef.Zoom
        End Select
        
        DiagramRef.PaperRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), .BackColor, BF
    End With
End Function

Private Sub Class_Initialize()
  Dim uGUID As GUID
  Dim sGUID As String
  Dim bGUID() As Byte
  Dim lLen As Long
  Dim RetVal As Long

  bGUID = String(40, 0)

  CoCreateGuid uGUID

  RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), 40)

  sGUID = bGUID
  If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
  Reference = Replace$(Replace$(Replace$(Left$(sGUID, RetVal), "-", ""), "{", ""), "}", "")
End Sub
