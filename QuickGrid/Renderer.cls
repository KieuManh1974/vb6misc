VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Renderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    Width As Long
    Height As Long
End Type

Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_NOPREFIX = &H800

Public mlTextHeight As Long
Public moTableInfo As TableInfo
Private moTheme As Theme
Public moDevice As Device
Public moTabStops As TabStops
Public moSelectionn As Selectionn
Public moVerticalScroll As Control
Public moHorizontalScroll As Control

Private mhBrush As Long
Private mhBrushEmpty As Long
Private mhBrushHighlightText As Long
Private mhBrushSelected As Long



Public Sub RenderTable()
    Dim lLineIndex As Long
    Dim vLine As Variant
    Dim lCellIndex As Long
    Dim rectArea As RECT
    Dim rectArea2 As RECT
    Dim sCell As String
    Dim lCharacterCount As Long
    Dim lCharacterCountEnd As Long
    Dim lOffsetY As Long
    Dim lHighlightStart As Long
    Dim lHighlightEnd As Long
    Dim lCellStartChar As Long
    Dim lCellEndChar As Long
    Dim bShowHighlight As Boolean
    Dim sizeDimensions2 As SIZE
    Dim sizeDimensions1 As SIZE
    Dim sColumnLetter As String
    Dim lStartLine As Long
    Dim lEndLine As Long

    lOffsetY = (moTableInfo.mlCellHeight - moTableInfo.mlCellSeparator - mlTextHeight) \ 2

    If moSelectionn.SelectionType = stText Then
        lHighlightStart = moSelectionn.StartPosition
        lHighlightEnd = moSelectionn.EndPosition
    End If
    
    lStartLine = moVerticalScroll.Value
    lEndLine = Int(moDevice.moDevice.Height / moTableInfo.mlCellHeight) - 1 + lStartLine
    
    If lEndLine >= moTableInfo.moTable.LastRow Then
        lEndLine = moTableInfo.moTable.LastRow
    End If

    moDevice.TopTableOffset(0) = -moVerticalScroll.Value * moTableInfo.mlCellHeight
    moDevice.LeftTableOffset(0) = -moTabStops.TabStop(moHorizontalScroll.Value)

    With moTableInfo
        rectArea.Top = 0
        rectArea.Left = 0
        rectArea.Bottom = .mlCellOffsetTop - .mlCellSeparator
        rectArea.Right = .mlCellOffsetLeft - .mlCellSeparator
    End With

    FillRect moDevice.hDC(0), rectArea, mhBrush
    
    moTabStops.Clear

    For lLineIndex = 0 To moTableInfo.moTable.LastRow
        For lCellIndex = 0 To UBound(moTableInfo.moTable.TableRow(lLineIndex))
            sCell = moTableInfo.moTable.TableCell(lLineIndex, lCellIndex)
            GetTextExtentPoint32 moDevice.hDC(0), sCell, Len(sCell), sizeDimensions1
            If sizeDimensions1.Width < moTabStops.DefaultWidth Then
                sizeDimensions1.Width = moTabStops.DefaultWidth
            End If
            If sizeDimensions1.Width > (moTabStops.TabStop(lCellIndex + 1) - moTabStops.TabStop(lCellIndex) - moTableInfo.mlCellSeparator) Then
                moTabStops.TabStop(lCellIndex + 1) = moTabStops.TabStop(lCellIndex) + sizeDimensions1.Width + moTableInfo.mlCellSeparator
            End If
        Next
    Next

    lCharacterCount = 0
    For lLineIndex = 0 To lStartLine - 1
        vLine = moTableInfo.moTable.TableRow(lLineIndex)
        For lCellIndex = 0 To UBound(vLine)
            sCell = vLine(lCellIndex)
            lCharacterCount = lCharacterCount + Len(sCell) + 1
        Next
        lCharacterCount = lCharacterCount + 1
    Next

    For lLineIndex = lStartLine To lEndLine
        vLine = moTableInfo.moTable.TableRow(lLineIndex)

        For lCellIndex = 0 To UBound(vLine)
            sCell = vLine(lCellIndex)
            lCharacterCountEnd = lCharacterCount + Len(sCell)
            
            With moTableInfo
                rectArea.Top = .mlCellOffsetTop + lLineIndex * .mlCellHeight + moDevice.TopTableOffset(0)
                rectArea.Left = .mlCellOffsetLeft + moTabStops.TabStop(lCellIndex) + moDevice.LeftTableOffset(0)
                rectArea.Bottom = .mlCellOffsetTop + (lLineIndex + 1) * .mlCellHeight - .mlCellSeparator + moDevice.TopTableOffset(0)
                rectArea.Right = .mlCellOffsetLeft + moTabStops.TabStop(lCellIndex + 1) - .mlCellSeparator + moDevice.LeftTableOffset(0)
            End With
            
            ' Draw box and text
            Select Case moSelectionn.SelectionType
                Case stTable
                    FillRect moDevice.hDC(0), rectArea, mhBrushSelected
                Case stCells
                    If moSelectionn.Cell(lLineIndex, lCellIndex) Then
                        FillRect moDevice.hDC(0), rectArea, mhBrushSelected
                    ElseIf sCell <> "" Then
                        FillRect moDevice.hDC(0), rectArea, mhBrush
                    Else
                        FillRect moDevice.hDC(0), rectArea, mhBrushEmpty
                    End If
                Case stRows
                    If moSelectionn.Row(lLineIndex) Then
                        FillRect moDevice.hDC(0), rectArea, mhBrushSelected
                    ElseIf sCell <> "" Then
                        FillRect moDevice.hDC(0), rectArea, mhBrush
                    Else
                        FillRect moDevice.hDC(0), rectArea, mhBrushEmpty
                    End If
                Case stColumns
                    If moSelectionn.Column(lCellIndex) Then
                        FillRect moDevice.hDC(0), rectArea, mhBrushSelected
                    ElseIf sCell <> "" Then
                        FillRect moDevice.hDC(0), rectArea, mhBrush
                    Else
                        FillRect moDevice.hDC(0), rectArea, mhBrushEmpty
                    End If
                Case Else
                    If sCell <> "" Then
                        FillRect moDevice.hDC(0), rectArea, mhBrush
                    Else
                        FillRect moDevice.hDC(0), rectArea, mhBrushEmpty
                    End If
            End Select
            SetTextColor moDevice.hDC(0), vbBlack
            DrawText moDevice.hDC(0), sCell, Len(sCell), rectArea, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
            
            If moSelectionn.SelectionType = stText Then
                bShowHighlight = True
                If lHighlightStart >= lCharacterCount And lHighlightStart <= lCharacterCountEnd Then
                    lCellStartChar = lHighlightStart - lCharacterCount
                ElseIf lHighlightStart < lCharacterCount Then
                    lCellStartChar = 0
                Else
                    bShowHighlight = False
                End If
                If lHighlightEnd >= lCharacterCount And lHighlightEnd <= lCharacterCountEnd Then
                    lCellEndChar = lHighlightEnd - lCharacterCount
                ElseIf lHighlightEnd > lCharacterCountEnd Then
                    lCellEndChar = lCharacterCountEnd - lCharacterCount
                Else
                    bShowHighlight = False
                End If

                If bShowHighlight Then
                    GetTextExtentPoint32 moDevice.hDC(0), Left$(sCell, lCellStartChar), lCellStartChar, sizeDimensions1
                    GetTextExtentPoint32 moDevice.hDC(0), Left$(sCell, lCellEndChar), lCellEndChar, sizeDimensions2
                    
                    If lHighlightEnd > lCharacterCountEnd Then
                        sizeDimensions2.Width = sizeDimensions2.Width + (rectArea.Right - rectArea.Left - sizeDimensions2.Width)
                    End If
                    
                    ' Draw highlight and text
                    With rectArea2
                        .Top = rectArea.Top + lOffsetY
                        .Bottom = .Top + mlTextHeight
                        .Left = rectArea.Left + sizeDimensions1.Width
                        .Right = rectArea.Left + sizeDimensions2.Width
                    End With
                    FillRect moDevice.hDC(0), rectArea2, mhBrushHighlightText
                    SetTextColor moDevice.hDC(0), vbWhite
                    DrawText moDevice.hDC(0), Mid$(sCell, lCellStartChar + 1, lCellEndChar - lCellStartChar), lCellEndChar - lCellStartChar, rectArea2, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
                End If
            End If

            If lCellIndex < UBound(vLine) Then
                lCharacterCount = lCharacterCountEnd + Len(moTableInfo.moTable.msColumnDelimiter)
            Else
                lCharacterCount = lCharacterCountEnd
            End If
        Next

        ' Draw row selector
        With moTableInfo
            rectArea.Top = .mlCellOffsetTop + lLineIndex * .mlCellHeight + moDevice.TopTableOffset(0)
            rectArea.Left = 0
            rectArea.Bottom = .mlCellOffsetTop + (lLineIndex + 1) * .mlCellHeight - .mlCellSeparator + moDevice.TopTableOffset(0)
            rectArea.Right = .mlCellOffsetLeft - .mlCellSeparator
        End With
            
        FillRect moDevice.hDC(0), rectArea, mhBrush ' row selector
        SetTextColor moDevice.hDC(0), vbBlack
        DrawText moDevice.hDC(0), CStr(lLineIndex + 1), Len(CStr(lLineIndex + 1)), rectArea, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_CENTER
        
        If lLineIndex < moTableInfo.moTable.LastRow Then
            lCharacterCount = lCharacterCountEnd + Len(moTableInfo.moTable.msRowDelimiter)
        Else
            lCharacterCount = lCharacterCountEnd
        End If
    Next
    
    ' Draw column selectors
    For lCellIndex = moHorizontalScroll.Value To moTableInfo.moTable.LastColumn
        With moTableInfo
            rectArea.Top = 0
            rectArea.Left = .mlCellOffsetLeft + moTabStops.TabStop(lCellIndex) + moDevice.LeftTableOffset(0)
            rectArea.Bottom = .mlCellOffsetTop - .mlCellSeparator
            rectArea.Right = .mlCellOffsetLeft + moTabStops.TabStop(lCellIndex + 1) - .mlCellSeparator + moDevice.LeftTableOffset(0)
        End With
        
        FillRect moDevice.hDC(0), rectArea, mhBrush
        sColumnLetter = ColumnLetter(lCellIndex)
        SetTextColor moDevice.hDC(0), vbBlack
        DrawText moDevice.hDC(0), sColumnLetter, Len(sColumnLetter), rectArea, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or DT_CENTER
    Next
    
    moVerticalScroll.Max = moTableInfo.moTable.LastRow
    moHorizontalScroll.Max = moTableInfo.moTable.LastColumn
End Sub

Public Sub Initialise()
    Dim sizeDimensions As SIZE
    
    GetTextExtentPoint32 moDevice.hDC(0), "X", 1, sizeDimensions
    mlTextHeight = sizeDimensions.Height
End Sub

Private Function ColumnLetter(ByVal lColumn As Long)
    If lColumn = 0 Then
        ColumnLetter = "A"
        Exit Function
    End If
    
    While lColumn > 0
        ColumnLetter = Chr$((lColumn Mod 26) + 65) & ColumnLetter
        lColumn = lColumn \ 26
    Wend
End Function


Public Property Set Theme(oTheme As Theme)
    Set moTheme = oTheme
    mhBrush = CreateSolidBrush(moTheme.mlCellColour)
    mhBrushEmpty = CreateSolidBrush(moTheme.mlEmptyCellColour)
    mhBrushSelected = CreateSolidBrush(moTheme.mlSelectedCellColour)
    mhBrushHighlightText = CreateSolidBrush(moTheme.mlSelectedText)
End Property

Private Sub Class_Terminate()
    DeleteObject mhBrush
    DeleteObject mhBrushHighlightText
    DeleteObject mhBrushEmpty
    DeleteObject mhBrushSelected
End Sub
