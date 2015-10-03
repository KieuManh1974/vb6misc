VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreaSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum HoverAreas
    haCell
    haColumnDivider
    haColumnSelector
    haRowSelector
    haTableSelector
    haOther
    haNone
End Enum

Public mlCellCharOffset As Long
Public mlTextCharOffset As Long
Public mlLength As Long

Public moCellPosition As New CellPosition

Public mhaAreaSelection As HoverAreas

Public moTableInfo As TableInfo
Public moTabStops As TabStops

Public Sub GetSelection(ByVal lX As Long, ByVal lY As Long)
    Dim lColumnDividePos As Long
    Dim lNextTabPosition As Long
    Dim lActualX As Long
    Dim lActualY As Long
    
    mhaAreaSelection = haOther
    
    If lX < moTableInfo.mlCellOffsetLeft And lY < moTableInfo.mlCellOffsetTop Then
        mhaAreaSelection = haTableSelector
    ElseIf lX < moTableInfo.mlCellOffsetLeft Then
        mhaAreaSelection = haRowSelector
        moCellPosition.RowIndex = (lY - moTableInfo.mlCellOffsetTop - moCellPosition.moDevice.TopTableOffset(0)) \ moTableInfo.mlCellHeight
        If moCellPosition.RowIndex > moTableInfo.moTable.LastRow Then
            mhaAreaSelection = haOther
        End If
    Else
        lActualX = lX - moTableInfo.mlCellOffsetLeft - moCellPosition.moDevice.LeftTableOffset(0)

        moCellPosition.ColumnIndex = moTabStops.GetTabStopForPosition(lActualX)
        lNextTabPosition = moTabStops.TabStop(moCellPosition.ColumnIndex + 1)
        
        If lNextTabPosition > 1 Then
            If lActualX >= (lNextTabPosition - moTableInfo.mlCellSeparator) And lActualX < lNextTabPosition Then
                mhaAreaSelection = haColumnDivider
                moCellPosition.ColumnIndex = moCellPosition.ColumnIndex + 1
                Exit Sub
            End If
        End If
        
        If lY < moTableInfo.mlCellOffsetTop Then
            mhaAreaSelection = haColumnSelector
            If moCellPosition.ColumnIndex > moTableInfo.moTable.LastColumn Then
                mhaAreaSelection = haOther
            End If
        Else
            lActualY = lY - moTableInfo.mlCellOffsetTop - moCellPosition.moDevice.TopTableOffset(0)
            
            mhaAreaSelection = haCell
            
            moCellPosition.RowIndex = lActualY \ moTableInfo.mlCellHeight

            If moCellPosition.ColumnIndex >= 0 Then
                moCellPosition.HorzPixelPosition = lActualX - moTabStops.TabStop(moCellPosition.ColumnIndex)
            Else
                moCellPosition.HorzPixelPosition = 0
            End If
            If moCellPosition.RowIndex > moTableInfo.moTable.LastRow Or moCellPosition.RowIndex < 0 Or moCellPosition.ColumnIndex < 0 Then
                mhaAreaSelection = haOther
            Else
                If moCellPosition.ColumnIndex > UBound(moTableInfo.moTable.TableRow(moCellPosition.RowIndex)) Then
                   mhaAreaSelection = haOther
                End If
            End If
        End If

    End If
    
    
End Sub
