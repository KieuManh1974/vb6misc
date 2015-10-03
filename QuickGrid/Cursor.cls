VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Private Type SIZE
    Width As Long
    Height As Long
End Type

Public moDevice As Device
Public moVerticalScroll As Control
Public moTableInfo As TableInfo
Public moPosition As CellPosition
Public mlTextHeight As Long
Public moTabStops As TabStops

Private mbCreated As Boolean

Public Sub MoveCursor()
    Dim lOffsetY As Long
    
    If Not mbCreated Then
        DestroyCaret
        CreateCaret moDevice.hWnd(0), 0&, 0&, mlTextHeight
        mbCreated = True
        ShowCaret moDevice.hWnd(0)
    End If
    
    lOffsetY = (moTableInfo.mlCellHeight - moTableInfo.mlCellSeparator - mlTextHeight) \ 2
    SetCaretPos moTabStops.TabStop(moPosition.ColumnIndex) + moTableInfo.mlCellOffsetLeft + moPosition.HorzPixelPosition + moDevice.LeftTableOffset(0), moPosition.RowIndex * moTableInfo.mlCellHeight + moTableInfo.mlCellOffsetTop + lOffsetY + moDevice.TopTableOffset(0)
    ShowCaret moDevice.hWnd(0)
End Sub

Public Sub HideCursor()
    HideCaret moDevice.hWnd(0)
End Sub

Public Sub RecreateCursor()
    mbCreated = False
    MoveCursor
End Sub

Public Sub Initialise()
    Dim sizeDimensions As SIZE
    
    GetTextExtentPoint32 moDevice.hDC(0), "X", 1, sizeDimensions
    mlTextHeight = sizeDimensions.Height
End Sub

