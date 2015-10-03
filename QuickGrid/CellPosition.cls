VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CellPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Private mlColumnIndex As Long
Private mlRowIndex As Long
Private mlHorzPixelPosition As Long
Private mlCellTextPosition As Long
Private mlTextPosition As Long

Public moTableInfo As TableInfo
Public moDevice As Device

Private Type SIZE
    Width As Long
    Height As Long
End Type

Private mlColumnIndexStamp As Long
Private mlRowIndexStamp As Long
Private mlHorzPixelPositionStamp As Long
Private mlCellTextPositionStamp As Long
Private mlTextPositionStamp As Long

Public Property Get ColumnIndex() As Long
    If mlTextPositionStamp > Largest(mlRowIndexStamp, mlColumnIndexStamp, mlCellTextPositionStamp, mlHorzPixelPositionStamp) Then
        ConvertTextPositionToHorzPixelPosition
    End If
    ColumnIndex = mlColumnIndex
End Property

Public Property Get RowIndex() As Long
    If mlTextPositionStamp > Largest(mlRowIndexStamp, mlColumnIndexStamp, mlCellTextPositionStamp, mlHorzPixelPositionStamp) Then
        ConvertTextPositionToHorzPixelPosition
    End If
    RowIndex = mlRowIndex
End Property

Public Property Get HorzPixelPosition() As Long
    If mlCellTextPositionStamp > mlHorzPixelPositionStamp Then
        ConvertCellTextPositionToTextPosition
    End If
    If mlTextPositionStamp > Largest(mlRowIndexStamp, mlColumnIndexStamp, mlHorzPixelPositionStamp) Then
        ConvertTextPositionToHorzPixelPosition
    End If

    HorzPixelPosition = mlHorzPixelPosition
End Property

Public Property Get CellTextPosition() As Long
    If Largest(mlColumnIndexStamp, mlRowIndexStamp, mlHorzPixelPositionStamp) > mlCellTextPositionStamp Then
        ConvertHorzPixelPositionToCellTextPosition
    End If
    CellTextPosition = mlCellTextPosition
End Property

Public Property Get TextPosition() As Long
    If Largest(mlColumnIndexStamp, mlRowIndexStamp, mlCellTextPositionStamp) > mlTextPositionStamp Then
        ConvertCellTextPositionToTextPosition
    End If
    TextPosition = mlTextPosition
End Property

Public Property Let ColumnIndex(ByVal lValue As Long)
    mlColumnIndexStamp = Largest(mlTextPositionStamp, mlCellTextPositionStamp) + 1
    mlColumnIndex = lValue
End Property

Public Property Let RowIndex(ByVal lValue As Long)
    mlRowIndexStamp = Largest(mlTextPositionStamp, mlCellTextPositionStamp) + 1
    mlRowIndex = lValue
End Property

Public Property Let HorzPixelPosition(ByVal lValue As Long)
    mlHorzPixelPositionStamp = mlTextPositionStamp + 1
    mlHorzPixelPosition = lValue
End Property

Public Property Let CellTextPosition(ByVal lValue As Long)
    mlCellTextPositionStamp = Largest(mlColumnIndexStamp, mlRowIndexStamp, mlHorzPixelPositionStamp) + 1
    mlCellTextPosition = lValue
End Property

Public Property Let TextPosition(ByVal lValue As Long)
    mlTextPositionStamp = Largest(mlColumnIndexStamp, mlRowIndexStamp, mlCellTextPositionStamp, mlHorzPixelPositionStamp) + 1
    mlTextPosition = lValue
End Property

Public Sub ConvertHorzPixelPositionToCellTextPosition()
    Dim lCharIndex As Long
    Dim sCellText As String
    Dim sizeDimensions As SIZE
    Dim lPreviousWidth As Long

    If UBound(moTableInfo.moTable.Table) = -1 Then
        RowIndex = 0
        ColumnIndex = 0
        CellTextPosition = 0
        TextPosition = 0
        HorzPixelPosition = 0
        mlRowIndexStamp = Largest(mlRowIndexStamp, mlColumnIndexStamp, mlCellTextPositionStamp, mlHorzPixelPositionStamp, mlTextPositionStamp)
        mlColumnIndexStamp = mlRowIndexStamp
        mlCellTextPositionStamp = mlRowIndexStamp
        mlHorzPixelPositionStamp = mlRowIndexStamp
        mlTextPositionStamp = mlRowIndexStamp
        Exit Sub
    End If
    
    If RowIndex <= moTableInfo.moTable.LastRow Then
        If UBound(moTableInfo.moTable.TableRow(RowIndex)) >= ColumnIndex Then
            sCellText = moTableInfo.moTable.TableCell(RowIndex, ColumnIndex)
        End If
    End If
    
    For lCharIndex = 1 To Len(sCellText)
        GetTextExtentPoint32 moDevice.hDC(0), Left$(sCellText, lCharIndex), lCharIndex, sizeDimensions

        If mlHorzPixelPosition < sizeDimensions.Width Then
            If mlHorzPixelPosition < (sizeDimensions.Width + lPreviousWidth) \ 2 Then
                mlHorzPixelPosition = lPreviousWidth
                mlCellTextPosition = lCharIndex - 1
                mlCellTextPositionStamp = mlHorzPixelPositionStamp
                Exit Sub
            Else
                mlHorzPixelPosition = sizeDimensions.Width
                mlCellTextPosition = lCharIndex
                mlCellTextPositionStamp = mlHorzPixelPositionStamp
                Exit Sub
            End If
        End If
        
        lPreviousWidth = sizeDimensions.Width
    Next
    mlHorzPixelPosition = lPreviousWidth
    mlCellTextPosition = lCharIndex - 1
    
    mlCellTextPositionStamp = mlHorzPixelPositionStamp
End Sub

Public Sub ConvertCellTextPositionToTextPosition()
    Dim lRowIndex As Long
    Dim lColumnIndex As Long
    Dim vRow As Variant
    Dim sRowText As String
    Dim vTextRows As Variant
    Dim sText As String
    
    If RowIndex > UBound(moTableInfo.moTable.Table) Then
        RowIndex = UBound(moTableInfo.moTable.Table)
    End If
    
    vTextRows = Array()
    For lRowIndex = 0 To mlRowIndex - 1
        sRowText = Join(moTableInfo.moTable.TableRow(lRowIndex), moTableInfo.moTable.msColumnDelimiter)
        ReDim Preserve vTextRows(UBound(vTextRows) + 1)
        
        vTextRows(UBound(vTextRows)) = sRowText
    Next
    If RowIndex > 0 Then
        sText = Join(vTextRows, moTableInfo.moTable.msRowDelimiter) & moTableInfo.moTable.msRowDelimiter
    End If
    
    If RowIndex > -1 Then
        If ColumnIndex > UBound(moTableInfo.moTable.TableRow(RowIndex)) Then
            ColumnIndex = UBound(moTableInfo.moTable.TableRow(mlRowIndex))
        End If
    Else
        ColumnIndex = 0
        TextPosition = 0
        Exit Sub
    End If
    
    For lColumnIndex = 0 To mlColumnIndex - 1
        If UBound(moTableInfo.moTable.TableRow(RowIndex)) > -1 Then
            sText = sText & moTableInfo.moTable.TableCell(RowIndex, lColumnIndex) & moTableInfo.moTable.msColumnDelimiter
        Else
            sText = sText & moTableInfo.moTable.msColumnDelimiter
        End If
    Next
    If UBound(moTableInfo.moTable.Table) > -1 Then
        If UBound(moTableInfo.moTable.TableRow(RowIndex)) > -1 Then
            sText = sText & Left$(moTableInfo.moTable.TableCell(RowIndex, ColumnIndex), CellTextPosition)
        End If
    End If
    mlTextPosition = Len(sText)
    
    mlTextPositionStamp = Largest(mlColumnIndexStamp, mlRowIndexStamp, mlCellTextPositionStamp)
End Sub

Public Sub ConvertTextPositionToHorzPixelPosition()
    Dim sPartialText As String
    Dim vLines As Variant
    Dim sCellText As String
    Dim sizeDimensions As SIZE
    
    If TextPosition = 0 Then
        ColumnIndex = 0
        RowIndex = 0
        HorzPixelPosition = 0
        Exit Sub
    End If
    
    sPartialText = Left$(moTableInfo.moTable.Text, TextPosition)
    vLines = Split(sPartialText, moTableInfo.moTable.msRowDelimiter)
    RowIndex = UBound(vLines)
    vLines = Split(vLines(UBound(vLines)), moTableInfo.moTable.msColumnDelimiter)
    mlColumnIndex = UBound(vLines)
    If mlColumnIndex = -1 Then
        HorzPixelPosition = 0
        ColumnIndex = 0
    Else
        sCellText = vLines(UBound(vLines))
        GetTextExtentPoint32 moDevice.hDC(0), sCellText, Len(sCellText), sizeDimensions
        HorzPixelPosition = sizeDimensions.Width
    End If
    
    mlHorzPixelPositionStamp = mlTextPositionStamp
    mlColumnIndexStamp = mlTextPositionStamp
    mlRowIndexStamp = mlTextPositionStamp
End Sub

Public Function Copy() As CellPosition
    Set Copy = New CellPosition
    With Copy
        .ColumnIndex = ColumnIndex
        .RowIndex = RowIndex
        If .ColumnIndex > -1 And .RowIndex > -1 Then
            .CellTextPosition = CellTextPosition
            .HorzPixelPosition = HorzPixelPosition
            '.TextPosition = TextPosition
        End If
        Set .moTableInfo = moTableInfo
        Set .moDevice = moDevice
    End With
End Function

Public Function CopyOf(oCellPosition) As CellPosition
    moDevice.hDC(0) = oCellPosition.moDevice.hDC(0)
    Set moTableInfo = oCellPosition.moTableInfo
    CellTextPosition = oCellPosition.CellTextPosition
    ColumnIndex = oCellPosition.ColumnIndex
    RowIndex = oCellPosition.RowIndex
    HorzPixelPosition = oCellPosition.HorzPixelPosition
    TextPosition = oCellPosition.TextPosition
End Function
