VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFilePath As String
Public Expanded As Boolean
Public Node As clsNode
Public Caption As String
Public Position As New clsPoint
Public Icon As IPictureDisp
Public Visible As Boolean

Private mlCaptionWidth As Long

Public Enum SelectionType
    stNone
    stExpand
    stIcon
    stVisibility
    stCaption
End Enum

Public Property Let FilePath(sFilePath As String)
    Dim oFile As File
    
    msFilePath = sFilePath
    Set Icon = GetIcon(sFilePath, 1)
    Caption = goFSO.GetBaseName(sFilePath)
    mlCaptionWidth = Form1.TextWidth(Caption)
End Property

Public Property Get FilePath() As String
    FilePath = msFilePath
End Property

Public Sub DrawTree(oCanvas As Object, bShowAll As Boolean)
    Dim lIndex As Long
    Dim pChildPosition As New clsPoint
    
    If Not bShowAll Then
        If Not Visible Then
            Exit Sub
        End If
    End If
    
    If Expanded Then
        oCanvas.PaintPicture goMinus, Position.X, Position.Y + 5
    Else
        oCanvas.PaintPicture goPlus, Position.X, Position.Y + 5
    End If
    

    If bShowAll Then
        If Visible Then
            oCanvas.PaintPicture goTicked, Position.X + 16, Position.Y + 3
        Else
            oCanvas.PaintPicture goUnticked, Position.X + 16, Position.Y + 3
        End If
        
        oCanvas.PaintPicture Icon, Position.X + 33, Position.Y
            
        oCanvas.CurrentX = Position.X + 52
        oCanvas.CurrentY = Position.Y + 2
        oCanvas.Print Caption
    Else
        oCanvas.PaintPicture Icon, Position.X + 16, Position.Y
        oCanvas.CurrentX = Position.X + 38
        oCanvas.CurrentY = Position.Y + 2
        oCanvas.Print Caption
    End If


'    pChildPosition.X = pPosition.X + 20
'    pChildPosition.Y = pPosition.Y + 20
'    For lIndex = 0 To Node.Count - 1
'        Node.ItemPhysical(lIndex).Value.DrawTree oCanvas, pChildPosition
'        pPosition.Y = pPosition.Y + 20
'    Next
End Sub

Public Function InRectangle(ByVal lXpos As Long, lYPos As Long) As SelectionType
    InRectangle = stNone
    If lXpos < Position.X Or lYPos < Position.Y Then
        Exit Function
    End If
    If lXpos > Position.X + 38 + mlCaptionWidth Then
        Exit Function
    End If
    If (lYPos - Position.Y) > 32 Then
        Exit Function
    End If
    
    If lXpos < (Position.X + 16) Then
        InRectangle = stExpand
    ElseIf lXpos < (Position.X + 38) Then
        InRectangle = stIcon
    Else
        InRectangle = stCaption
    End If
End Function
