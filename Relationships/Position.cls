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

Public CanvasRef As Canvas
Public ParserRef As IParseObject

Private Const Margin As Single = 1

Public Function LeftTextPosition()
    LeftTextPosition = PosX - CanvasRef.TextWidth(Name) / 2
End Function

Public Function RightTextPosition()
    RightTextPosition = CanvasRef.TextWidth(Name) / 2 + PosX
End Function

Public Function TopTextPosition()
    TopTextPosition = PosY - CanvasRef.TextHeight(Name) / 2
End Function

Public Function BottomTextPosition()
    BottomTextPosition = PosY + CanvasRef.TextHeight(Name) / 2
End Function

Public Function Height()
    Height = CanvasRef.TextHeight(Name)
End Function

Public Function RenderName(Optional iColour As Long = vbRed)
    With CanvasRef
        .FillColor = iColour
        .ForeColor = iColour
        CanvasRef.Circle (.sTopLeftX + LeftTextPosition - Margin, .sTopLeftY + PosY), Margin + Height / 2 - 1
        CanvasRef.Circle (.sTopLeftX + RightTextPosition + Margin, .sTopLeftY + PosY), Margin + Height / 2 - 1
        
        CanvasRef.Line (.sTopLeftX + LeftTextPosition - Margin, .sTopLeftY + TopTextPosition - Margin)-(.sTopLeftX + RightTextPosition + Margin, .sTopLeftY + BottomTextPosition + Margin), iColour, BF
        .CurrentX = .sTopLeftX + LeftTextPosition
        .CurrentY = .sTopLeftY + TopTextPosition
        
        .ForeColor = vbWhite
        CanvasRef.Print Name
    End With
End Function

Public Function ClearName()
    With CanvasRef
        .FillColor = .BackColor
        .ForeColor = .BackColor
        CanvasRef.Circle (.sTopLeftX + LeftTextPosition - Margin, .sTopLeftY + PosY), Margin + Height / 2 - 1
        CanvasRef.Circle (.sTopLeftX + RightTextPosition + Margin, .sTopLeftY + PosY), Margin + Height / 2 - 1
        
        CanvasRef.Line (.sTopLeftX + LeftTextPosition - Margin, .sTopLeftY + TopTextPosition - Margin)-(.sTopLeftX + RightTextPosition + Margin, .sTopLeftY + BottomTextPosition + Margin), .BackColor, BF
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
