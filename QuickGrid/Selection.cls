VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Selectionn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SelectionTypes
    stNone
    stText
    stColumns
    stRows
    stCells
    stTable
End Enum

Private moPositionA As Long
Private moPositionB As Long
Private mvSet As Variant
Private mbInitialised As Boolean
Private mstSelectionType As SelectionTypes

Public moTable As Table
Public moCursor As Cursor

Public Property Let SelectionType(ByVal stSelectionType As SelectionTypes)
    If stSelectionType <> mstSelectionType Then
        mstSelectionType = stSelectionType
        moPositionB = moPositionA
        mvSet = Array(Array(False))
        mbInitialised = False
        Select Case stSelectionType
            Case SelectionTypes.stText
                moCursor.RecreateCursor
            Case Else
                moCursor.HideCursor
        End Select
    End If
End Property

Public Property Get SelectionType() As SelectionTypes
    SelectionType = mstSelectionType
End Property

Public Property Let StartPosition(lPosition As Long)
    moPositionA = lPosition
End Property

Public Property Let EndPosition(lPosition As Long)
    moPositionB = lPosition
End Property

'Public Property Let Position(lPosition As Long)
'    If mstSelectionType <> stText Then
'        moPositionA = lPosition
'    End If
'    moPositionB = lPosition
'End Property

Public Property Get StartPosition() As Long
    If moPositionB < moPositionA Then
        StartPosition = moPositionB
    Else
        StartPosition = moPositionA
    End If
End Property

Public Property Get EndPosition() As Long
    If moPositionA < moPositionB Then
        EndPosition = moPositionB
    Else
        EndPosition = moPositionA
    End If
End Property

Public Property Let Column(lColumn As Long, bSwitch As Boolean)
    If Not mbInitialised Then
        ReDim mvSet(moTable.LastColumn)
        mbInitialised = True
    End If
    mvSet(lColumn) = bSwitch
End Property

Public Property Let Row(lRow As Long, bSwitch As Boolean)
    If Not mbInitialised Then
        ReDim mvSet(moTable.LastRow)
        mbInitialised = True
    End If
    mvSet(lRow) = bSwitch
End Property

Public Property Get Column(lColumn As Long) As Boolean
    If Not mbInitialised Then
        ReDim mvSet(moTable.LastColumn)
        mbInitialised = True
    End If
    Column = mvSet(lColumn)
End Property

Public Property Get Row(lRow As Long) As Boolean
    If Not mbInitialised Then
        ReDim mvSet(moTable.LastRow)
        mbInitialised = True
    End If
    Row = mvSet(lRow)
End Property

Public Property Let Cell(lRow As Long, lColumn As Long, bSwitch As Boolean)
    If Not mbInitialised Then
        BuildCells
        mbInitialised = True
    End If
    mvSet(lRow)(lColumn) = bSwitch
End Property

Public Property Get Cell(lRow As Long, lColumn As Long) As Boolean
    If Not mbInitialised Then
        BuildCells
        mbInitialised = True
    End If
    If lRow <= UBound(mvSet) Then
        If lColumn <= UBound(mvSet(lRow)) Then
            Cell = mvSet(lRow)(lColumn)
        End If
    End If
End Property

Public Property Let Range(lStartRow As Long, lStartColumn As Long, lEndRow As Long, lEndColumn As Long, bSwitch As Boolean)
    Dim lRow As Long
    Dim lColumn As Long
    
    BuildCells
    
    For lRow = lStartRow To lEndRow
        For lColumn = lStartColumn To lEndColumn
            mvSet(lRow)(lColumn) = bSwitch
        Next
    Next
End Property

Public Property Let ColumnRange(lStartColumn As Long, lEndColumn As Long, bSwitch As Boolean)
    Dim lColumn As Long
    Dim lStep As Long
    
    lStep = 1
    If lStartColumn > lEndColumn Then
        lStep = -1
    End If
    
    ReDim mvSet(moTable.LastColumn)
    mbInitialised = True

    For lColumn = lStartColumn To lEndColumn Step lStep
        mvSet(lColumn) = bSwitch
    Next
End Property

Public Property Let RowRange(lStartRow As Long, lEndRow As Long, bSwitch As Boolean)
    Dim lRow As Long
    Dim lStep As Long
    
    lStep = 1
    If lStartRow > lEndRow Then
        lStep = -1
    End If
    
    ReDim mvSet(moTable.LastRow)
    mbInitialised = True
    
    For lRow = lStartRow To lEndRow Step lStep
        mvSet(lRow) = bSwitch
    Next
End Property

Private Sub BuildCells()
    Dim lHeight As Long
    Dim lRow As Long
    Dim lWidth As Long
    
    lHeight = moTable.LastRow
    lWidth = moTable.LastColumn
    
    ReDim mvSet(lHeight) As Variant
    For lRow = 0 To lHeight
        mvSet(lRow) = PaddedArray(lWidth, False)
    Next
End Sub

Private Sub Class_Initialize()
    mvSet = Array(Array(False))
    mstSelectionType = stNone
End Sub
