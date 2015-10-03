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
Public PosX As Single
Public PosY As Single
Public Colour As Long
Public Orientation As Long

Public DiagramRef As Diagram
Public ParserRef As IParseObject

Private Const Margin As Single = 1

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

Public Function Radius()
    Radius = Margin + Height / 2 - 1
End Function

Public Function RenderName(Optional iColour As Long = -1)
    Dim nTextX As Single
    Dim nTextY As Single
    
    With DiagramRef
        If iColour = -1 Then
            iColour = Colour
        End If
        .FillColor = iColour
        .FillStyle = 0
        If Name <> "" Then
            .DrawWidth = 5
            .ForeColor = .BackColor
            DiagramRef.Circle (.sTopLeftX + PosX, .sTopLeftY + PosY), Radius + 2
            .DrawWidth = 1
        Else
            .DrawWidth = 1
            .ForeColor = iColour
            DiagramRef.Circle (.sTopLeftX + PosX, .sTopLeftY + PosY), Radius
            .DrawWidth = 1
        End If

        If Trim$(Name) <> "" Then
            Select Case Orientation
                Case 0, 4
                    nTextY = .sTopLeftY + PosY - Height / 2
                Case 1, 2, 3
                    nTextY = .sTopLeftY + PosY + 10
                Case 5, 6, 7
                    nTextY = .sTopLeftY + PosY - Height - 10
            End Select
            
            Select Case Orientation
                Case 7, 0, 1
                    nTextX = .sTopLeftX + PosX + 10
                Case 2, 6
                    nTextX = .sTopLeftX + PosX - .TextWidth(Name) / 2
                Case 3, 4, 5
                    nTextX = .sTopLeftX + PosX - .TextWidth(Name) - 10
            End Select
            
            DiagramRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), .BackColor, BF
            .CurrentX = nTextX
            .CurrentY = nTextY
            
            .ForeColor = vbWhite
            DiagramRef.Print Name
        End If
    End With
End Function

Public Function ClearName()

    Dim nTextX As Single
    Dim nTextY As Single
    
    With DiagramRef
        .FillColor = .BackColor
        .ForeColor = .BackColor
        DiagramRef.Circle (.sTopLeftX + PosX, .sTopLeftY + PosY), Radius
        
        Select Case Orientation
            Case 0, 4
                nTextY = .sTopLeftY + PosY + Height / 2
            Case 1, 2, 3
                nTextY = .sTopLeftY + PosY + 10
            Case 5, 6, 7
                nTextY = .sTopLeftY + PosY - Height - 10
        End Select
        
        Select Case Orientation
            Case 7, 0, 1
                nTextX = .sTopLeftX + PosX + 10
            Case 2, 6
                nTextX = .sTopLeftX + PosX - .TextWidth(Name) / 2
            Case 3, 4, 5
                nTextX = .sTopLeftX + PosX - .TextWidth(Name) - 10
        End Select
        
        DiagramRef.Line (nTextX, nTextY)-Step(.TextWidth(Name), Height), .BackColor, BF
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
