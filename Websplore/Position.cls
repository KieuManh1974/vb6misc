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

Private Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32.dll" (lpLogPen As LOGPEN) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
 
Private Type POINTAPI
   X As Long
   Y As Long
End Type
 
Private Type LOGPEN
   lopnStyle As Long
   lopnWidth As POINTAPI
   lopnColor As Long
End Type
 
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6
Private Const PS_SOLID = 0

Public Reference As String

Public Name As String
Public Path As String
Public Icon As IPictureDisp

Public PosX As Single
Public PosY As Single
Public Orientation As Long

Public Colour As Long
Public Shape As Long
Public Radius As Single

Public DiagramRef As Diagram
Public ParserRef As IParseObject

Private Const Margin As Single = 1
Public Highlighted As Boolean

Public Function Highlight(ByVal bThisHighlight As Boolean)
    Highlighted = bThisHighlight
    RenderName
End Function

Public Function LeftTextPosition()
    LeftTextPosition = PosX - DiagramRef.TextWidth(Name) / 2
End Function

Public Function RightTextPosition()
    RightTextPosition = DiagramRef.TextWidth(Name) / 2 + PosX
End Function

Public Function TopTextPosition()
    TopTextPosition = PosY - DiagramRef.TextHeight(Name) / 2
End Function

Public Function BottomTextPosition()
    BottomTextPosition = PosY + DiagramRef.TextHeight(Name) / 2
End Function

Public Function Height()
    Height = DiagramRef.TextHeight(Name)
End Function

Public Function RenderName(Optional iColour As Long = -1)
    Dim nTextX As Single
    Dim nTextY As Single
    Dim oFSO As FileSystemObject
    
    With DiagramRef
        If Not Highlighted Then
            If iColour = -1 Then
                iColour = Colour
            End If
        Else
            iColour = &HFFC0C0
        End If
        
        .FillColor = iColour
        .FillStyle = 0
        If Name <> "" Then
            If Icon Is Nothing Then
                Set oFSO = New FileSystemObject
                If oFSO.FileExists(Path) Or oFSO.FolderExists(Path) Then
                    Set Icon = GetIcon(Path, 1)
                End If
            End If
            
            If Not Icon Is Nothing Then
                DiagramRef.PaintPicture Icon, .nTopLeftX + PosX - 8, .nTopLeftY + PosY - 8
            Else
                .DrawWidth = 5
                .ForeColor = .BackColor
                DrawShape iColour
                .DrawWidth = 1
            End If
        Else
            .DrawWidth = 1
            .ForeColor = iColour
            DiagramRef.Circle (.nTopLeftX + PosX, .nTopLeftY + PosY), Radius

            .DrawWidth = 1
        End If

        Select Case Orientation
            Case 0, 4
                nTextY = .nTopLeftY + PosY - Height / 2
            Case 1, 2, 3
                nTextY = .nTopLeftY + PosY + 10
            Case 5, 6, 7
                nTextY = .nTopLeftY + PosY - Height - 10
        End Select
        
        Select Case Orientation
            Case 7, 0, 1
                nTextX = .nTopLeftX + PosX + 12
            Case 2, 6
                nTextX = .nTopLeftX + PosX - .TextWidth(Name) / 2
            Case 3, 4, 5
                nTextX = .nTopLeftX + PosX - .TextWidth(Name) - 12
        End Select
        
        DiagramRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), IIf(Highlighted, &HFFC0C0, .BackColor), BF
        .CurrentX = nTextX
        .CurrentY = nTextY
        
        .ForeColor = vbBlack
        DiagramRef.Print Name
    End With
End Function

Public Function ClearName()
    Dim nTextX As Single
    Dim nTextY As Single
    
    With DiagramRef
        .FillColor = .BackColor
        .ForeColor = .BackColor
        DiagramRef.Circle (.nTopLeftX + PosX, .nTopLeftY + PosY), Radius
        
        Select Case Orientation
            Case 0, 4
                nTextY = .nTopLeftY + PosY - Height / 2
            Case 1, 2, 3
                nTextY = .nTopLeftY + PosY + 10
            Case 5, 6, 7
                nTextY = .nTopLeftY + PosY - Height - 10
        End Select
        
        Select Case Orientation
            Case 7, 0, 1
                nTextX = .nTopLeftX + PosX + 12
            Case 2, 6
                nTextX = .nTopLeftX + PosX - .TextWidth(Name) / 2
            Case 3, 4, 5
                nTextX = .nTopLeftX + PosX - .TextWidth(Name) - 12
        End Select
        
        DiagramRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), .BackColor, BF
    End With
End Function

Public Function Snap()
    PosX = Int(PosX / 20) * 20
    PosY = Int(PosY / 20) * 20
End Function
    
Private Sub Class_Initialize()
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim Retval As Long
    
    Radius = 7
    Shape = 0
    bGUID = String(40, 0)
    CoCreateGuid uGUID
    Retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), 40)
    sGUID = bGUID
    If (Asc(Mid$(sGUID, Retval, 1)) = 0) Then Retval = Retval - 1
    Reference = Replace$(Replace$(Replace$(Left$(sGUID, Retval), "-", ""), "{", ""), "}", "")
End Sub

Private Sub DrawShape(ByVal lColour As Long)
    Dim nAngle As Single
    Dim nStep As Single
    Dim pi2 As Single
    
    Dim lRetval As Long
    Dim hPen As Long
    Dim hOldPen As Long
    Dim tPenInfo As LOGPEN
    Dim tPoints(10) As POINTAPI
    Dim lIndex As Long
    
    pi2 = Atn(1) * 8
    With DiagramRef
        Select Case Shape
            Case 0
                DiagramRef.Circle (.nTopLeftX + PosX, .nTopLeftY + PosY), Radius
            Case 1, 2
            Case Else
                With tPenInfo
                   .lopnColor = lColour
                   .lopnStyle = PS_SOLID
                End With
                hPen = CreatePenIndirect(tPenInfo)
                hOldPen = SelectObject(DiagramRef.hdc, hPen)
                
                lIndex = 0
                For nStep = 0 To Shape
                    nAngle = nStep * pi2 / CSng(Shape)
                    tPoints(lIndex).X = Cos(-pi2 / 4 - nAngle) * Radius + PosX + .nTopLeftX
                    tPoints(lIndex).Y = Sin(-pi2 / 4 - nAngle) * Radius + PosY + .nTopLeftY
                    lIndex = lIndex + 1
                Next
                lRetval = PolyPolygon(DiagramRef.hdc, tPoints(0), Shape, 1)
                Call SelectObject(DiagramRef.hdc, hOldPen)
                DeleteObject hPen
        End Select
    End With
End Sub


